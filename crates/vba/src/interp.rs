//! The VBA interpreter: runs procedures over the parsed modules, calling out to the host for its
//! object model (cells, MsgBox, …). Clean-room, from [MS-VBAL].
//!
//! The interpreter owns no host state; everything host-specific goes through [`Host`]. Recursion
//! and loop iterations are bounded so a runaway macro can't hang GridCraft.

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::ast::{Arg, CaseTest, ExitKind, Expr, Line, Module, OnError, Param, ProcKind, Procedure, ResumeKind, Stmt, UnOp, VarDecl};
use crate::builtins::{self, Collection, Dictionary};
use crate::error::{VbaError, err};
use crate::value::{self, BinOp, HostRef, Obj, Udt, Value};

/// How deep calls may nest, and how many statements a single run may execute.
const MAX_CALL_DEPTH: usize = 256;
const MAX_STEPS: u64 = 500_000_000;
/// Statement blocks and expressions being run at once, across calls: past it the run stops with
/// "Out of stack space" instead of exhausting the thread's stack.
const MAX_NEST: usize = 3000;
/// Native stack, in bytes, a run may use below its first block: past it the run stops with "Out
/// of stack space". Leaves room for the host's own work (a formula it calculates has its own
/// budget) on the macro thread.
const STACK_BUDGET: usize = 512 * 1024;

/// The address of a local variable, which shows how far down the native stack this call is.
#[inline(always)]
fn stack_position() -> usize {
    let marker = 0u8;
    std::hint::black_box(&marker) as *const u8 as usize
}

/// What the interpreter asks of its host (the Excel object model). Every method may fail with a
/// `VbaError`; `Ok(None)` from the lookups means "not something I provide".
pub trait Host {
    /// A global name (`Range`, `Cells`, `ActiveSheet`, `Application`, `vbOKOnly`, `xlUp`…):
    /// a value, or a host object to send members to. `Ok(None)` → not a host global.
    fn global(&mut self, name: &str) -> Result<Option<Value>, VbaError>;
    /// Get `obj.member(args)` (a property or method). `set` distinguishes `x.P = v` targets.
    fn get_member(&mut self, obj: &HostRef, member: &str, args: &[Value]) -> Result<Value, VbaError>;
    /// Let/Set `obj.member(args) = value`.
    fn set_member(&mut self, obj: &HostRef, member: &str, args: &[Value], value: Value, by_ref: bool) -> Result<(), VbaError>;
    /// Call a host global procedure form like `MsgBox`/`InputBox` when it isn't a plain global.
    fn call_global(&mut self, _name: &str, _args: &[Value]) -> Result<Option<Value>, VbaError> {
        Ok(None)
    }
    /// `For Each` over a host object: its items as a Vec (snapshot).
    fn enumerate(&mut self, _obj: &HostRef) -> Result<Vec<Value>, VbaError> {
        Err(err(438, "Object doesn't support this property or method"))
    }
    /// `CreateObject("Scripting.Dictionary")` and friends; `Ok(None)` → unknown class.
    fn create_object(&mut self, _class: &str) -> Result<Option<Value>, VbaError> {
        Ok(None)
    }
    /// Show a message box; returns the button code. Default: no UI, returns OK (1).
    fn msgbox(&mut self, _prompt: &str, _buttons: i64, _title: &str) -> Result<i64, VbaError> {
        Ok(1)
    }
    /// Ask for a line of input; `None` = cancelled.
    fn inputbox(&mut self, _prompt: &str, _title: &str, _default: &str) -> Result<Option<String>, VbaError> {
        Ok(None)
    }
    /// `Debug.Print` output.
    fn debug_print(&mut self, _text: &str) {}
    /// The text of a file a macro opens for input (`Open … For Input`).
    fn read_text_file(&mut self, _path: &str) -> Result<String, VbaError> {
        Err(err(75, "Path/File access error"))
    }
    /// Writes a file a macro closes (`Open … For Output | Append`): the whole text, or added to
    /// the end with `append`.
    fn write_text_file(&mut self, _path: &str, _text: &str, _append: bool) -> Result<(), VbaError> {
        Err(err(75, "Path/File access error"))
    }
}

/// A class-module instance (`New Thing`).
#[derive(Debug)]
pub struct Instance {
    pub class: String,
    pub fields: HashMap<String, Value>,
}

/// A variable slot, shared so `ByRef` arguments alias the caller's storage.
type Slot = Rc<RefCell<Value>>;

#[derive(Clone)]
enum Flow {
    Normal,
    Exit(ExitKind),
    GoTo(String),
    End,
}

struct Scope {
    vars: HashMap<String, Slot>,
    /// `Static` locals, kept across calls (keyed by procedure).
    statics: HashMap<String, Slot>,
}

/// The parsed project the interpreter runs.
pub struct Program {
    pub modules: Vec<Module>,
    /// Module name of each module (lower-cased) → index.
    index: HashMap<String, usize>,
}

impl Program {
    pub fn new(modules: Vec<Module>) -> Program {
        let index = modules.iter().enumerate().map(|(i, m)| (m.name.to_lowercase(), i)).collect();
        Program { modules, index }
    }

    /// Parses and loads modules from their (name, source) pairs.
    pub fn parse(sources: &[(String, String)]) -> Result<Program, VbaError> {
        let mut modules = vec![];
        for (name, src) in sources {
            modules.push(crate::parser::parse_module(name, src)?);
        }
        Ok(Program::new(modules))
    }

    /// Runnable macros: public Subs with no required arguments, as `Module.Proc`.
    pub fn macros(&self) -> Vec<String> {
        let mut out = vec![];
        for m in &self.modules {
            for p in &m.procs {
                if p.kind == ProcKind::Sub && p.public && p.params.iter().all(|a| a.optional || a.param_array) {
                    out.push(format!("{}.{}", m.name, p.name));
                }
            }
        }
        out
    }

    /// Every procedure, public or not, as `Module.Proc`.
    pub fn procedures(&self) -> Vec<String> {
        self.modules.iter().flat_map(|m| m.procs.iter().map(move |p| format!("{}.{}", m.name, p.name))).collect()
    }

    /// Every public and private procedure, as (module index, procedure index).
    fn find_proc(&self, name: &str, from_module: Option<usize>) -> Option<(usize, usize)> {
        // A module's own procedures first, then any public one.
        if let Some(mi) = from_module
            && let Some(pi) = self.modules.get(mi).and_then(|m| m.procs.iter().position(|p| p.name.eq_ignore_ascii_case(name) && !is_set_let(p)))
        {
            return Some((mi, pi));
        }
        for (mi, m) in self.modules.iter().enumerate() {
            if let Some(pi) = m.procs.iter().position(|p| p.name.eq_ignore_ascii_case(name) && p.public && !is_set_let(p)) {
                return Some((mi, pi));
            }
        }
        None
    }
}

fn is_set_let(p: &Procedure) -> bool {
    matches!(p.kind, ProcKind::PropertyLet | ProcKind::PropertySet)
}

pub struct Interp<'h> {
    pub program: Program,
    pub host: &'h mut dyn Host,
    globals: HashMap<String, Slot>,
    consts: HashMap<String, Value>,
    /// Enum member name (lower) → value.
    enums: HashMap<String, i64>,
    /// Type name (lower) → field list, for building UDT defaults.
    types: HashMap<String, Vec<VarDecl>>,
    depth: usize,
    /// Blocks and expressions being run (see [`MAX_NEST`]).
    nest: usize,
    /// Where on the native stack the outermost block started (see [`STACK_BUDGET`]).
    stack_top: usize,
    stack_budget: usize,
    steps: u64,
    /// The current `With` object stack.
    with_stack: Vec<Value>,
    text_compare: bool,
    /// The last error raised, for `Err`.
    last_error: Option<VbaError>,
    err_cleared: bool,
    /// Set from another thread to stop the macro (the Stop button).
    cancel: Option<std::sync::Arc<std::sync::atomic::AtomicBool>>,
    /// Files opened with `Open`.
    files: crate::files::Files,
}

impl<'h> Interp<'h> {
    pub fn new(program: Program, host: &'h mut dyn Host) -> Interp<'h> {
        let mut it = Interp {
            program,
            host,
            globals: HashMap::new(),
            consts: HashMap::new(),
            enums: HashMap::new(),
            types: HashMap::new(),
            depth: 0,
            nest: 0,
            stack_top: 0,
            stack_budget: STACK_BUDGET,
            steps: 0,
            with_stack: vec![],
            text_compare: false,
            last_error: None,
            err_cleared: true,
            cancel: None,
            files: crate::files::Files::default(),
        };
        it.load_module_level();
        it
    }

    fn load_module_level(&mut self) {
        let modules = std::mem::take(&mut self.program.modules);
        for m in &modules {
            self.text_compare |= m.compare_text;
            for (name, items) in &m.enums {
                let _ = name;
                let mut next = 0i64;
                for (k, e) in items {
                    let v = match e {
                        Some(expr) => self.const_expr(expr).unwrap_or(next),
                        None => next,
                    };
                    self.enums.insert(k.to_lowercase(), v);
                    next = v + 1;
                }
            }
            for (name, fields) in &m.types {
                self.types.insert(name.to_lowercase(), fields.clone());
            }
        }
        // Constants can refer to earlier ones and to enum members.
        for m in &modules {
            for (name, _, expr, _) in &m.consts {
                if let Ok(v) = self.eval_const(expr) {
                    self.consts.insert(name.to_lowercase(), v);
                }
            }
        }
        // Module-level variables start Empty (or their type's default); populated lazily.
        for m in &modules {
            for (decl, _) in &m.vars {
                let v = self.default_for(decl);
                self.globals.entry(decl.name.to_lowercase()).or_insert_with(|| Rc::new(RefCell::new(v)));
            }
        }
        self.program.modules = modules;
    }

    fn const_expr(&self, e: &Expr) -> Option<i64> {
        match e {
            Expr::Lit(crate::lexer::Tok::Int(n, _)) => Some(*n),
            Expr::Unary(UnOp::Neg, inner) => self.const_expr(inner).map(|n| -n),
            Expr::Name(n) => self.enums.get(&n.to_lowercase()).copied(),
            _ => None,
        }
    }

    fn eval_const(&mut self, e: &Expr) -> Result<Value, VbaError> {
        // Constant expressions can't call the host; evaluate in an empty scope.
        let mut scope = Scope { vars: HashMap::new(), statics: HashMap::new() };
        self.eval(e, &mut scope)
    }

    fn default_for(&self, decl: &VarDecl) -> Value {
        if let Some(dims) = &decl.dims {
            if dims.is_empty() {
                return Value::Empty; // dynamic, un-dimensioned
            }
            // Fixed array: computed when run (needs evaluation); start Empty, ReDim-like at use.
            return Value::Empty;
        }
        self.type_default(decl.ty.name.as_deref())
    }

    fn type_default(&self, ty: Option<&str>) -> Value {
        match ty.map(|t| t.to_ascii_lowercase()).as_deref() {
            Some("long" | "integer" | "byte" | "long long" | "longlong") => Value::Long(0),
            Some("double" | "single" | "currency") => Value::Double(0.0),
            Some("boolean") => Value::Bool(false),
            Some("string") => Value::str(""),
            Some("date") => Value::Date(0.0),
            Some("variant") | None => Value::Empty,
            Some("object" | "collection" | "range" | "worksheet" | "workbook") => Value::Nothing,
            Some(other) => {
                if let Some(fields) = self.types.get(other) {
                    Value::Udt(Box::new(Udt {
                        type_name: other.into(),
                        fields: fields.iter().map(|f| (f.name.clone(), self.type_default(f.ty.name.as_deref()))).collect(),
                    }))
                } else {
                    Value::Empty
                }
            }
        }
    }

    // ------------------------------------------------------------ public entry

    /// Runs the named procedure (`Module.Proc` or `Proc`) with `args`, returning its result.
    pub fn run(&mut self, name: &str, args: Vec<Value>) -> Result<Value, VbaError> {
        self.steps = 0;
        let (module, proc) = match name.split_once('.') {
            Some((m, p)) => (self.program.index.get(&m.to_lowercase()).copied(), p.to_string()),
            None => (None, name.to_string()),
        };
        let (mi, pi) = self.program.find_proc(&proc, module).ok_or_else(|| err(35, format!("Sub or Function not defined: {name}")))?;
        let slots: Vec<Slot> = args.into_iter().map(|v| Rc::new(RefCell::new(v))).collect();
        let r = self.call_proc(mi, pi, slots, false);
        // Files left open are closed (and written) when the macro ends, as VBA does on `End`.
        let closed = self.close_files(&[]);
        let r = r?;
        closed?;
        Ok(r)
    }

    /// Closes files `numbers` (every open file when empty), writing those opened for output.
    fn close_files(&mut self, numbers: &[i64]) -> Result<(), VbaError> {
        let numbers = if numbers.is_empty() { self.files.numbers() } else { numbers.to_vec() };
        let mut first_error = None;
        for n in numbers {
            let r = self.files.close(n).and_then(|w| match w {
                Some((path, text, append)) => self.host.write_text_file(&path, &text, append),
                None => Ok(()),
            });
            if let Err(e) = r {
                first_error.get_or_insert(e);
            }
        }
        first_error.map_or(Ok(()), Err)
    }

    /// The names of the runnable macros (public Subs with no required arguments).
    pub fn macros(&self) -> Vec<String> {
        self.program.macros()
    }

    // ------------------------------------------------------------ calls

    /// Lets the run use `bytes` of native stack (default [`STACK_BUDGET`]): for a host that runs
    /// macros on a thread with a large stack of its own.
    pub fn set_stack_budget(&mut self, bytes: usize) {
        self.stack_budget = bytes;
    }

    /// Stops the run when `flag` becomes true (checked every few hundred statements).
    pub fn set_cancel(&mut self, flag: std::sync::Arc<std::sync::atomic::AtomicBool>) {
        self.cancel = Some(flag);
    }

    fn step(&mut self) -> Result<(), VbaError> {
        self.steps += 1;
        if self.steps > MAX_STEPS {
            return Err(err(28, "The macro ran too long and was stopped"));
        }
        if self.steps.is_multiple_of(256)
            && let Some(c) = &self.cancel
            && c.load(std::sync::atomic::Ordering::Relaxed)
        {
            // Not trappable by On Error: number -2 unwinds like End.
            return Err(VbaError {
                number: -2,
                description: "Code execution has been interrupted".into(),
                source: String::new(),
                line: None,
                module: None,
            });
        }
        Ok(())
    }

    fn call_proc(&mut self, mi: usize, pi: usize, args: Vec<Slot>, want_value: bool) -> Result<Value, VbaError> {
        self.depth += 1;
        if self.depth > MAX_CALL_DEPTH {
            self.depth -= 1;
            return Err(err(28, "Out of stack space"));
        }
        let proc = match self.program.modules.get(mi).and_then(|m| m.procs.get(pi)) {
            Some(p) => p.clone(),
            None => {
                self.depth -= 1;
                return Err(err(35, "Sub or Function not defined"));
            }
        };
        let r = self.call_proc_inner(mi, &proc, args, want_value);
        self.depth -= 1;
        r
    }

    fn call_proc_inner(&mut self, _mi: usize, proc: &Procedure, mut args: Vec<Slot>, _want_value: bool) -> Result<Value, VbaError> {
        let mut scope = Scope { vars: HashMap::new(), statics: HashMap::new() };
        // Bind parameters.
        for (ai, param) in proc.params.iter().enumerate() {
            if param.param_array {
                let rest: Vec<Value> = args.get(ai..).unwrap_or(&[]).iter().map(|s| s.borrow().clone()).collect();
                scope.vars.insert(param.name.to_lowercase(), Rc::new(RefCell::new(builtins::array_1d(rest))));
                break;
            }
            let slot = match args.get_mut(ai) {
                Some(s) if !matches!(&*s.borrow(), Value::Empty if is_missing(param)) => {
                    if param.by_val {
                        Rc::new(RefCell::new(coerce(&s.borrow(), param.ty.name.as_deref())?))
                    } else {
                        s.clone()
                    }
                }
                _ => {
                    let v = match &param.default {
                        Some(e) => self.eval(e, &mut scope)?,
                        None if param.optional => {
                            if param.ty.name.is_none() {
                                Value::Error(448)
                            } else {
                                self.type_default(param.ty.name.as_deref())
                            }
                        }
                        None => return Err(err(449, format!("Argument not optional: {}", param.name))),
                    };
                    Rc::new(RefCell::new(v))
                }
            };
            scope.vars.insert(param.name.to_lowercase(), slot);
        }
        // The function's return slot is a variable named like the procedure.
        let ret_default = self.type_default(proc.returns.name.as_deref());
        let returns_value = matches!(proc.kind, ProcKind::Function | ProcKind::PropertyGet);
        if returns_value {
            scope.vars.insert(proc.name.to_lowercase(), Rc::new(RefCell::new(ret_default.clone())));
        }
        let mut ctx = ProcCtx { name: proc.name.clone(), handler: None, in_handler: false };
        let flow = self.exec_lines(&proc.body, &mut scope, &mut ctx)?;
        let result = match flow {
            Flow::End => return Err(VbaError { number: -1, description: "End".into(), source: ctx.name, line: None, module: None }),
            _ if returns_value => scope.vars.get(&proc.name.to_lowercase()).map(|s| s.borrow().clone()).unwrap_or(ret_default),
            _ => Value::Empty,
        };
        Ok(result)
    }

    // ------------------------------------------------------------ statements

    /// Runs `f` one block or expression deeper.
    fn nested<T>(&mut self, f: impl FnOnce(&mut Self) -> Result<T, VbaError>) -> Result<T, VbaError> {
        let here = stack_position();
        if self.nest == 0 {
            self.stack_top = here;
        }
        if self.nest >= MAX_NEST || self.stack_top.abs_diff(here) > self.stack_budget {
            return Err(err(28, "Out of stack space"));
        }
        self.nest += 1;
        let r = f(self);
        self.nest -= 1;
        r
    }

    fn exec_lines(&mut self, lines: &[Line], scope: &mut Scope, ctx: &mut ProcCtx) -> Result<Flow, VbaError> {
        self.nested(|it| it.exec_lines_inner(lines, scope, ctx))
    }

    fn exec_lines_inner(&mut self, lines: &[Line], scope: &mut Scope, ctx: &mut ProcCtx) -> Result<Flow, VbaError> {
        let mut i = 0usize;
        while i < lines.len() {
            self.step()?;
            match &lines[i] {
                Line::Label(_) => {
                    i += 1;
                    continue;
                }
                Line::Stmt(stmt, line) => {
                    let flow = self.exec_stmt(stmt, scope, ctx, *line);
                    let flow = match flow {
                        Ok(f) => f,
                        Err(e) => self.handle_error(e, ctx, *line)?,
                    };
                    match flow {
                        Flow::Normal => i += 1,
                        Flow::GoTo(label) => {
                            match find_label(lines, &label) {
                                Some(j) => {
                                    i = j;
                                    ctx.in_handler = ctx.handler.as_deref() == Some(label.as_str());
                                }
                                // The label is elsewhere in the procedure (another block); bubble up.
                                None => return Ok(Flow::GoTo(label)),
                            }
                        }
                        other => return Ok(other),
                    }
                }
            }
        }
        Ok(Flow::Normal)
    }

    /// Turns a raised error into control flow per the active `On Error` setting.
    fn handle_error(&mut self, e: VbaError, ctx: &mut ProcCtx, line: usize) -> Result<Flow, VbaError> {
        if e.number < 0 {
            return Err(e); // End / internal signal
        }
        let mut e = e;
        e.line = Some(line);
        self.last_error = Some(e.clone());
        self.err_cleared = false;
        if ctx.handler.as_deref() == Some("\u{0}next") {
            return Ok(Flow::Normal);
        }
        match &ctx.handler {
            Some(label) if !ctx.in_handler => Ok(Flow::GoTo(label.clone())),
            _ => Err(e),
        }
    }

    fn exec_stmt(&mut self, stmt: &Stmt, scope: &mut Scope, ctx: &mut ProcCtx, line: usize) -> Result<Flow, VbaError> {
        match stmt {
            Stmt::Dim { decls, is_static } => {
                for d in decls {
                    let v = self.dim_value(d, scope)?;
                    let key = d.name.to_lowercase();
                    if *is_static {
                        let slot = ctx_static(scope, &ctx.name, &key, || v.clone());
                        scope.vars.insert(key, slot);
                    } else {
                        scope.vars.insert(key, Rc::new(RefCell::new(v)));
                    }
                }
                Ok(Flow::Normal)
            }
            Stmt::Const(list) => {
                for (n, _, e) in list {
                    let v = self.eval(e, scope)?;
                    scope.vars.insert(n.to_lowercase(), Rc::new(RefCell::new(v)));
                }
                Ok(Flow::Normal)
            }
            Stmt::ReDim { preserve, items } => {
                for (target, bounds, ty) in items {
                    let name = match target {
                        Expr::Name(n) => n.clone(),
                        _ => return Err(err(13, "Type mismatch")),
                    };
                    let b = self.eval_bounds(bounds, scope)?;
                    let fill = self.type_default(ty.name.as_deref());
                    let new = value::Array::new(b, fill)?;
                    let slot = self.slot_for(&name, scope);
                    if *preserve {
                        let old = slot.borrow().clone();
                        let merged = preserve_array(old, new)?;
                        *slot.borrow_mut() = merged;
                    } else {
                        *slot.borrow_mut() = Value::Array(Rc::new(RefCell::new(new)));
                    }
                }
                Ok(Flow::Normal)
            }
            Stmt::Assign { target, value } => {
                let v = self.eval_value(value, scope)?;
                self.assign(target, v, scope, false)?;
                Ok(Flow::Normal)
            }
            Stmt::Set { target, value } => {
                let v = self.eval(value, scope)?;
                self.assign(target, v, scope, true)?;
                Ok(Flow::Normal)
            }
            Stmt::Call(e) => {
                self.eval_call_stmt(e, scope)?;
                Ok(Flow::Normal)
            }
            Stmt::If { branches, otherwise } => {
                for (cond, body) in branches {
                    if self.eval_value(cond, scope)?.to_bool()? {
                        return self.exec_lines(body, scope, ctx);
                    }
                }
                if let Some(body) = otherwise {
                    return self.exec_lines(body, scope, ctx);
                }
                Ok(Flow::Normal)
            }
            Stmt::Select { subject, cases, otherwise } => {
                let subj = self.eval_value(subject, scope)?;
                for (tests, body) in cases {
                    if self.case_matches(&subj, tests, scope)? {
                        return self.exec_lines(body, scope, ctx);
                    }
                }
                if let Some(body) = otherwise {
                    return self.exec_lines(body, scope, ctx);
                }
                Ok(Flow::Normal)
            }
            Stmt::For { var, from, to, step, body } => self.exec_for(ForHead { var, from, to, step: step.as_ref() }, body, scope, ctx),
            Stmt::ForEach { var, collection, body } => self.exec_for_each(var, collection, body, scope, ctx),
            Stmt::Do { pre, post, body } => self.exec_do(pre.as_ref(), post.as_ref(), body, scope, ctx),
            Stmt::With { object, body } => {
                let obj = self.eval_ref(object, scope)?;
                self.with_stack.push(obj);
                let flow = self.exec_lines(body, scope, ctx);
                self.with_stack.pop();
                flow
            }
            Stmt::Exit(kind) => Ok(Flow::Exit(*kind)),
            Stmt::GoTo(label) => Ok(Flow::GoTo(label.clone())),
            Stmt::GoSub(_) | Stmt::Return => Err(err(5, "GoSub is not supported")),
            Stmt::OnError(mode) => {
                ctx.handler = match mode {
                    OnError::GoTo(l) => Some(l.clone()),
                    OnError::ResumeNext => Some("\u{0}next".into()),
                    OnError::Zero => None,
                };
                ctx.in_handler = false;
                Ok(Flow::Normal)
            }
            Stmt::Resume(kind) => {
                ctx.in_handler = false;
                match kind {
                    ResumeKind::Next | ResumeKind::Here => {
                        // We can't re-run the exact failing line without its position in the tree;
                        // Resume Next continues after the handler returns. Approximate by Normal.
                        Ok(Flow::Normal)
                    }
                    ResumeKind::Label(l) => Ok(Flow::GoTo(l.clone())),
                }
            }
            Stmt::Erase(targets) => {
                for t in targets {
                    if let Expr::Name(n) = t {
                        let slot = self.slot_for(n, scope);
                        let cleared = match &*slot.borrow() {
                            Value::Array(a) => {
                                let b = a.borrow();
                                if b.fixed {
                                    let fill = b.items.first().cloned().unwrap_or(Value::Empty);
                                    Value::Array(Rc::new(RefCell::new(value::Array {
                                        bounds: b.bounds.clone(),
                                        items: vec![fill; b.items.len()],
                                        fixed: true,
                                    })))
                                } else {
                                    Value::Empty
                                }
                            }
                            _ => Value::Empty,
                        };
                        *slot.borrow_mut() = cleared;
                    }
                }
                Ok(Flow::Normal)
            }
            Stmt::DebugPrint(items) => {
                let mut out = String::new();
                for it in items {
                    out.push_str(&self.eval_value(it, scope)?.to_str().unwrap_or_default());
                }
                self.host.debug_print(&out);
                Ok(Flow::Normal)
            }
            Stmt::End => Ok(Flow::End),
            Stmt::Stop => Ok(Flow::Normal),
            Stmt::RaiseEvent(..) => Ok(Flow::Normal),
            Stmt::Open { path, mode, number } => {
                let path = self.eval_value(path, scope)?.to_str()?;
                let n = crate::files::number(&self.eval_value(number, scope)?)?;
                if self.files.is_open(n) {
                    return Err(err(55, "File already open"));
                }
                let text = match mode {
                    crate::files::FileMode::Input => {
                        let t = self.host.read_text_file(&path)?;
                        value::check_length(t.len())?;
                        Some(t)
                    }
                    // Output starts the file empty now, as VBA does.
                    crate::files::FileMode::Output => {
                        self.host.write_text_file(&path, "", false)?;
                        None
                    }
                    crate::files::FileMode::Append => None,
                };
                self.files.open(n, &path, *mode, text)?;
                Ok(Flow::Normal)
            }
            Stmt::Close(numbers) => {
                let mut ns = vec![];
                for e in numbers {
                    ns.push(crate::files::number(&self.eval_value(e, scope)?)?);
                }
                self.close_files(&ns)?;
                Ok(Flow::Normal)
            }
            Stmt::PrintFile { number, items } => {
                let n = crate::files::number(&self.eval_value(number, scope)?)?;
                let mut values = Vec::with_capacity(items.len());
                for (e, sep) in items {
                    values.push((self.eval_value(e, scope)?, *sep));
                }
                self.files.print(n, &values)?;
                Ok(Flow::Normal)
            }
            Stmt::WriteFile { number, items } => {
                let n = crate::files::number(&self.eval_value(number, scope)?)?;
                let mut values = Vec::with_capacity(items.len());
                for e in items {
                    values.push(self.eval_value(e, scope)?);
                }
                self.files.write(n, &values)?;
                Ok(Flow::Normal)
            }
            Stmt::InputFile { number, targets, line } => {
                let n = crate::files::number(&self.eval_value(number, scope)?)?;
                for t in targets {
                    let v = if *line { Value::str(self.files.line_input(n)?) } else { self.files.input_field(n)? };
                    self.assign(t, v, scope, false)?;
                }
                Ok(Flow::Normal)
            }
            Stmt::Unsupported(what) => Err(err(5, format!("The macro uses `{what}`, which GridCraft can't run yet"))),
        }
        .map_err(|mut e| {
            if e.line.is_none() {
                e.line = Some(line);
            }
            e
        })
    }

    fn exec_for(&mut self, head: ForHead<'_>, body: &[Line], scope: &mut Scope, ctx: &mut ProcCtx) -> Result<Flow, VbaError> {
        let ForHead { var, from, to, step } = head;
        let start = self.eval_value(from, scope)?.to_f64()?;
        let end = self.eval_value(to, scope)?.to_f64()?;
        let step_v = match step {
            Some(e) => self.eval_value(e, scope)?.to_f64()?,
            None => 1.0,
        };
        if step_v == 0.0 && start != end {
            return Err(err(5, "For step cannot be zero"));
        }
        let mut cur = start;
        loop {
            self.step()?;
            if (step_v > 0.0 && cur > end) || (step_v < 0.0 && cur < end) {
                break;
            }
            self.assign(var, Value::from_f64_natural(cur), scope, false)?;
            match self.exec_lines(body, scope, ctx)? {
                Flow::Normal => {}
                Flow::Exit(ExitKind::For) => break,
                other => return Ok(other),
            }
            // Read the control variable back (the body may have changed it).
            cur = self.eval(var, scope)?.to_f64()? + step_v;
        }
        Ok(Flow::Normal)
    }

    fn exec_for_each(&mut self, var: &Expr, collection: &Expr, body: &[Line], scope: &mut Scope, ctx: &mut ProcCtx) -> Result<Flow, VbaError> {
        let coll = self.eval_ref(collection, scope)?;
        let items = self.iterate(&coll)?;
        for item in items {
            self.step()?;
            self.assign(var, item, scope, true)?;
            match self.exec_lines(body, scope, ctx)? {
                Flow::Normal => {}
                Flow::Exit(ExitKind::For) => break,
                other => return Ok(other),
            }
        }
        Ok(Flow::Normal)
    }

    fn iterate(&mut self, coll: &Value) -> Result<Vec<Value>, VbaError> {
        match coll {
            Value::Array(a) => Ok(a.borrow().items.clone()),
            Value::Object(Obj::Collection(c)) => Ok(c.borrow().items.clone()),
            Value::Object(Obj::Dictionary(d)) => Ok(d.borrow().pairs.iter().map(|(k, _)| k.clone()).collect()),
            Value::Object(Obj::Host(h)) => self.host.enumerate(h),
            Value::Nothing | Value::Empty => Ok(vec![]),
            _ => Err(err(438, "Object doesn't support this property or method")),
        }
    }

    fn exec_do(
        &mut self,
        pre: Option<&(bool, Expr)>,
        post: Option<&(bool, Expr)>,
        body: &[Line],
        scope: &mut Scope,
        ctx: &mut ProcCtx,
    ) -> Result<Flow, VbaError> {
        loop {
            self.step()?;
            if let Some((while_cond, e)) = pre {
                let c = self.eval_value(e, scope)?.to_bool()?;
                if c != *while_cond {
                    break;
                }
            }
            match self.exec_lines(body, scope, ctx)? {
                Flow::Normal => {}
                Flow::Exit(ExitKind::Do) => break,
                other => return Ok(other),
            }
            if let Some((while_cond, e)) = post {
                let c = self.eval_value(e, scope)?.to_bool()?;
                if c != *while_cond {
                    break;
                }
            }
        }
        Ok(Flow::Normal)
    }

    fn case_matches(&mut self, subj: &Value, tests: &[CaseTest], scope: &mut Scope) -> Result<bool, VbaError> {
        for t in tests {
            let hit = match t {
                CaseTest::Value(e) => value::compare(subj, &self.eval_value(e, scope)?, self.text_compare)? == std::cmp::Ordering::Equal,
                CaseTest::Range(a, b) => {
                    let lo = self.eval_value(a, scope)?;
                    let hi = self.eval_value(b, scope)?;
                    value::compare(subj, &lo, self.text_compare)? != std::cmp::Ordering::Less
                        && value::compare(subj, &hi, self.text_compare)? != std::cmp::Ordering::Greater
                }
                CaseTest::Is(op, e) => {
                    let r = value::binary(*op, subj, &self.eval_value(e, scope)?, self.text_compare)?;
                    r.to_bool()?
                }
            };
            if hit {
                return Ok(true);
            }
        }
        Ok(false)
    }

    fn eval_bounds(&mut self, bounds: &[(Option<Expr>, Expr)], scope: &mut Scope) -> Result<Vec<(i64, i64)>, VbaError> {
        let base = self.option_base();
        let mut out = vec![];
        for (lo, hi) in bounds {
            let lo = match lo {
                Some(e) => self.eval(e, scope)?.to_i64()?,
                None => base,
            };
            out.push((lo, self.eval(hi, scope)?.to_i64()?));
        }
        Ok(out)
    }

    fn option_base(&self) -> i64 {
        self.program.modules.first().map(|m| m.option_base).unwrap_or(0)
    }

    fn dim_value(&mut self, d: &VarDecl, scope: &mut Scope) -> Result<Value, VbaError> {
        if let Some(dims) = &d.dims {
            if dims.is_empty() {
                return Ok(Value::Empty);
            }
            let b = self.eval_bounds(dims, scope)?;
            let fill = self.type_default(d.ty.name.as_deref());
            let mut arr = value::Array::new(b, fill)?;
            arr.fixed = true;
            return Ok(Value::Array(Rc::new(RefCell::new(arr))));
        }
        if d.ty.new
            && let Some(cls) = &d.ty.name
        {
            return self.new_object(cls);
        }
        Ok(self.type_default(d.ty.name.as_deref()))
    }

    // ------------------------------------------------------------ expressions

    fn eval(&mut self, e: &Expr, scope: &mut Scope) -> Result<Value, VbaError> {
        self.nested(|it| it.eval_inner(e, scope))
    }

    fn eval_inner(&mut self, e: &Expr, scope: &mut Scope) -> Result<Value, VbaError> {
        self.step()?;
        match e {
            Expr::Lit(tok) => Ok(lit_value(tok)),
            Expr::Me => self.host.global("Me")?.ok_or_else(|| err(91, "Object variable or With block variable not set")),
            Expr::Missing => Ok(Value::Error(448)),
            Expr::Name(n) => self.eval_name(n, &[], scope),
            Expr::WithMember(m) => {
                let obj = self.with_stack.last().cloned().ok_or_else(|| err(91, "Invalid use of a With member"))?;
                self.member(&obj, m, &[], false)
            }
            Expr::Member(obj, m) => {
                let o = self.eval_ref(obj, scope)?;
                self.member(&o, m, &[], false)
            }
            Expr::Bang(obj, key) => {
                let o = if matches!(**obj, Expr::Name(ref n) if n == "\u{0}with") {
                    self.with_stack.last().cloned().ok_or_else(|| err(91, "With"))?
                } else {
                    self.eval_ref(obj, scope)?
                };
                self.member(&o, "Item", &[Value::str(key.as_str())], false)
            }
            Expr::Unary(UnOp::Neg, inner) => value::negate(&self.eval_value(inner, scope)?),
            Expr::Unary(UnOp::Not, inner) => value::not(&self.eval_value(inner, scope)?),
            Expr::Binary(op, a, b) => self.eval_binary(*op, a, b, scope),
            Expr::New(cls) => self.new_object(cls),
            Expr::TypeOf(e, cls) => {
                let v = self.eval_ref(e, scope)?;
                Ok(Value::Bool(v.type_name().eq_ignore_ascii_case(cls) || (cls.eq_ignore_ascii_case("Object") && v.is_object())))
            }
            Expr::Call(target, args) => self.eval_call(target, args, scope, false),
        }
    }

    /// Like `eval`, but returns objects by reference (not their default property).
    fn eval_ref(&mut self, e: &Expr, scope: &mut Scope) -> Result<Value, VbaError> {
        match e {
            Expr::Call(t, args) if matches!(**t, Expr::Name(ref n) if n == "\u{0}paren") => self.eval_ref(&args[0].value, scope),
            _ => self.eval(e, scope),
        }
    }

    /// An object used where a value is expected gives its default property (`Range("A1")` in
    /// `x = Range("A1")` or `Len(Range("A1"))` is the cell's value).
    fn deref(&mut self, v: Value) -> Result<Value, VbaError> {
        match v {
            Value::Object(Obj::Host(h)) => self.host.get_member(&h, "_Default", &[]),
            Value::Object(Obj::Collection(_) | Obj::Dictionary(_) | Obj::Instance(_)) => {
                Err(err(450, "Wrong number of arguments or invalid property assignment"))
            }
            other => Ok(other),
        }
    }

    /// `eval`, then the default property of an object.
    fn eval_value(&mut self, e: &Expr, scope: &mut Scope) -> Result<Value, VbaError> {
        let v = self.eval(e, scope)?;
        self.deref(v)
    }

    fn eval_binary(&mut self, op: BinOp, a: &Expr, b: &Expr, scope: &mut Scope) -> Result<Value, VbaError> {
        if op == BinOp::Is {
            let x = self.eval_ref(a, scope)?;
            let y = self.eval_ref(b, scope)?;
            return value::binary(BinOp::Is, &x, &y, false);
        }
        let x = self.eval_value(a, scope)?;
        // Short-circuit is not VBA semantics; And/Or evaluate both sides.
        let y = self.eval_value(b, scope)?;
        value::binary(op, &x, &y, self.text_compare)
    }

    fn eval_name(&mut self, name: &str, args: &[Value], scope: &mut Scope) -> Result<Value, VbaError> {
        let key = name.to_lowercase();
        if key == "\u{0}paren" {
            return Ok(args.first().cloned().unwrap_or(Value::Empty));
        }
        // 0) Literal keywords.
        if args.is_empty() {
            match key.as_str() {
                "true" => return Ok(Value::Bool(true)),
                "false" => return Ok(Value::Bool(false)),
                "nothing" => return Ok(Value::Nothing),
                "null" => return Ok(Value::Null),
                "empty" => return Ok(Value::Empty),
                _ => {}
            }
        }
        // 1) Local or module variable (with indices, that's an array/default-member access).
        if let Some(slot) = scope.vars.get(&key).or_else(|| self.globals.get(&key)).cloned() {
            let v = slot.borrow().clone();
            if args.is_empty() {
                return Ok(v);
            }
            return self.index_or_default(&v, args);
        }
        // 2) Constant / enum.
        if args.is_empty() {
            if let Some(v) = self.consts.get(&key) {
                return Ok(v.clone());
            }
            if let Some(v) = self.enums.get(&key) {
                return Ok(Value::Long(*v as i32));
            }
        }
        // 3) Built-in function.
        if let Some(v) = self.call_builtin(name, args, scope)? {
            return Ok(v);
        }
        // 4) User procedure.
        if let Some((mi, pi)) = self.program.find_proc(name, None) {
            let slots: Vec<Slot> = args.iter().map(|v| Rc::new(RefCell::new(v.clone()))).collect();
            return self.call_proc(mi, pi, slots, true);
        }
        if key == "err" && args.is_empty() {
            return Ok(Value::Object(Obj::Err));
        }
        // 5) Host global or host procedure.
        if let Some(v) = self.host.global(name)? {
            if args.is_empty() {
                return Ok(v);
            }
            return self.index_or_default(&v, args);
        }
        if let Some(v) = self.host.call_global(name, args)? {
            return Ok(v);
        }
        Err(err(if args.is_empty() { 2003 } else { 35 }, format!("Variable or procedure not defined: {name}")))
    }

    fn call_builtin(&mut self, name: &str, args: &[Value], _scope: &mut Scope) -> Result<Option<Value>, VbaError> {
        let upper = name.to_ascii_uppercase();
        // Built-ins work on values: objects give their default property, except where the
        // function is about the object itself.
        let keeps_objects = matches!(
            upper.as_str(),
            "TYPENAME"
                | "ISOBJECT"
                | "VARTYPE"
                | "ISNULL"
                | "ISEMPTY"
                | "ARRAY"
                | "IIF"
                | "CHOOSE"
                | "SWITCH"
                | "ISARRAY"
                | "UBOUND"
                | "LBOUND"
                | "JOIN"
                | "FILTER"
        );
        let owned: Vec<Value>;
        let args: &[Value] = if keeps_objects || !args.iter().any(|a| matches!(a, Value::Object(Obj::Host(_)))) {
            args
        } else {
            let mut v = Vec::with_capacity(args.len());
            for a in args {
                v.push(self.deref(a.clone())?);
            }
            owned = v;
            &owned
        };
        match upper.as_str() {
            "MSGBOX" => {
                let prompt = args.first().map(|v| v.to_str()).transpose()?.unwrap_or_default();
                let buttons = args.get(1).map(|v| v.to_i64()).transpose()?.unwrap_or(0);
                let title = args.get(2).map(|v| v.to_str()).transpose()?.unwrap_or_else(|| "GridCraft".into());
                return Ok(Some(Value::from_i64(self.host.msgbox(&prompt, buttons, &title)?)));
            }
            "INPUTBOX" => {
                let prompt = args.first().map(|v| v.to_str()).transpose()?.unwrap_or_default();
                let title = args.get(1).map(|v| v.to_str()).transpose()?.unwrap_or_else(|| "GridCraft".into());
                let default = args.get(2).map(|v| v.to_str()).transpose()?.unwrap_or_default();
                return Ok(Some(match self.host.inputbox(&prompt, &title, &default)? {
                    Some(s) => Value::str(s),
                    None => Value::str(""),
                }));
            }
            "CREATEOBJECT" | "GETOBJECT" => {
                let class = args.first().map(|v| v.to_str()).transpose()?.unwrap_or_default();
                if let Some(v) = self.create_builtin_object(&class)? {
                    return Ok(Some(v));
                }
                return self.host.create_object(&class).map(|o| o.or(Some(Value::Nothing)));
            }
            "DOEVENTS" => return Ok(Some(Value::Integer(0))),
            "EOF" => return Ok(Some(Value::Bool(self.files.eof(crate::files::number(&arg0(args))?)?))),
            "LOF" => return Ok(Some(Value::from_i64(self.files.lof(crate::files::number(&arg0(args))?)?))),
            "FREEFILE" => return Ok(Some(Value::from_i64(self.files.free()?))),
            "ARRAY" => return Ok(Some(builtins::array_1d(args.to_vec()))),
            _ => {}
        }
        builtins::call(name, args, self.text_compare)
    }

    fn create_builtin_object(&self, class: &str) -> Result<Option<Value>, VbaError> {
        let c = class.to_ascii_lowercase();
        if c.ends_with("dictionary") {
            return Ok(Some(Value::Object(Obj::Dictionary(Rc::new(RefCell::new(Dictionary::default()))))));
        }
        if c == "collection" {
            return Ok(Some(Value::Object(Obj::Collection(Rc::new(RefCell::new(Collection::default()))))));
        }
        Ok(None)
    }

    fn new_object(&mut self, cls: &str) -> Result<Value, VbaError> {
        if cls.eq_ignore_ascii_case("Collection") {
            return Ok(Value::Object(Obj::Collection(Rc::new(RefCell::new(Collection::default())))));
        }
        // A class module.
        if let Some(mi) = self.program.modules.iter().position(|m| m.name.eq_ignore_ascii_case(cls)) {
            let fields = self
                .program
                .modules
                .get(mi)
                .map(|m| m.vars.iter().map(|(d, _)| (d.name.to_lowercase(), self.type_default(d.ty.name.as_deref()))).collect())
                .unwrap_or_default();
            return Ok(Value::Object(Obj::Instance(Rc::new(RefCell::new(Instance { class: cls.to_string(), fields })))));
        }
        if let Some(v) = self.host.create_object(cls)? {
            return Ok(v);
        }
        Err(err(429, format!("ActiveX component can't create object: {cls}")))
    }

    fn eval_call(&mut self, target: &Expr, args: &[Arg], scope: &mut Scope, _stmt: bool) -> Result<Value, VbaError> {
        // `(expr)` grouping.
        if let Expr::Name(n) = target
            && n == "\u{0}paren"
        {
            return self.eval(&args[0].value, scope);
        }
        if let Expr::Name(n) = target
            && let Some((mi, pi)) = self.program.find_proc(n, None)
            && !self.local_is_indexable(n, scope)
        {
            let slots = self.arg_slots(mi, pi, args, scope)?;
            return self.call_proc(mi, pi, slots, true);
        }
        let argv = self.eval_args(args, scope)?;
        match target {
            Expr::Name(n) => self.eval_name(n, &argv, scope),
            Expr::WithMember(m) => {
                let obj = self.with_stack.last().cloned().ok_or_else(|| err(91, "With"))?;
                self.member(&obj, m, &argv, false)
            }
            Expr::Member(obj, m) => {
                let o = self.eval_ref(obj, scope)?;
                self.member(&o, m, &argv, false)
            }
            Expr::Bang(obj, key) => {
                let o = self.eval_ref(obj, scope)?;
                let mut a = vec![Value::str(key.as_str())];
                a.extend(argv);
                self.member(&o, "Item", &a, false)
            }
            // `f(x)(y)`: call the result.
            inner => {
                let v = self.eval(inner, scope)?;
                self.index_or_default(&v, &argv)
            }
        }
    }

    /// Whether a local or module variable with this name can be indexed/called (an array, object
    /// or string) — so `Name(args)` is an index, not a procedure call.
    fn local_is_indexable(&self, name: &str, scope: &Scope) -> bool {
        let key = name.to_lowercase();
        scope
            .vars
            .get(&key)
            .or_else(|| self.globals.get(&key))
            .is_some_and(|s| matches!(&*s.borrow(), Value::Array(_) | Value::Object(_) | Value::String(_)))
    }

    fn arg_slots(&mut self, mi: usize, pi: usize, args: &[Arg], scope: &mut Scope) -> Result<Vec<Slot>, VbaError> {
        let params = self.program.modules.get(mi).and_then(|m| m.procs.get(pi)).map(|p| p.params.clone()).unwrap_or_default();
        let has_named = args.iter().any(|a| a.name.is_some());
        let mut slots: Vec<Option<Slot>> = vec![None; params.len().max(args.len())];
        let mut pos = 0usize;
        for a in args {
            let idx = match &a.name {
                Some(n) => {
                    params.iter().position(|p| p.name.eq_ignore_ascii_case(n)).ok_or_else(|| err(448, format!("Named argument not found: {n}")))?
                }
                None => {
                    let i = pos;
                    pos += 1;
                    i
                }
            };
            if idx >= slots.len() {
                slots.resize(idx + 1, None);
            }
            slots[idx] = Some(self.one_arg_slot(&a.value, scope, a.by_val)?);
        }
        let last = slots.iter().rposition(|x| x.is_some()).map(|i| i + 1).unwrap_or(0);
        let n = if has_named { params.len().max(last) } else { last };
        Ok(slots.into_iter().take(n).map(|x| x.unwrap_or_else(|| Rc::new(RefCell::new(Value::Empty)))).collect())
    }

    fn one_arg_slot(&mut self, e: &Expr, scope: &mut Scope, by_val: bool) -> Result<Slot, VbaError> {
        if matches!(e, Expr::Missing) {
            return Ok(Rc::new(RefCell::new(Value::Error(448))));
        }
        let parenthesised = matches!(e, Expr::Call(t, _) if matches!(&**t, Expr::Name(n) if n == "\u{0}paren"));
        if !by_val
            && !parenthesised
            && let Expr::Name(n) = e
        {
            let key = n.to_lowercase();
            if let Some(sl) = scope.vars.get(&key).or_else(|| self.globals.get(&key)) {
                return Ok(sl.clone());
            }
        }
        Ok(Rc::new(RefCell::new(self.eval_ref(e, scope)?)))
    }

    fn eval_args(&mut self, args: &[Arg], scope: &mut Scope) -> Result<Vec<Value>, VbaError> {
        // Named arguments aren't resolved for host calls (the host maps them); for user calls we
        // reorder below. Here we keep order and pass missing as a sentinel.
        let mut out = vec![];
        for a in args {
            if matches!(a.value, Expr::Missing) {
                out.push(Value::Error(448));
            } else {
                out.push(self.eval_ref(&a.value, scope)?);
            }
        }
        Ok(out)
    }

    /// `v(args)`: an array index, a Collection/Dictionary item, a default member, or a host call.
    fn index_or_default(&mut self, v: &Value, args: &[Value]) -> Result<Value, VbaError> {
        match v {
            Value::Array(a) => {
                let idx: Result<Vec<i64>, _> = args.iter().map(|x| x.to_i64()).collect();
                let a = a.borrow();
                Ok(a.items.get(a.index(&idx?)?).cloned().unwrap_or(Value::Empty))
            }
            Value::Object(Obj::Collection(c)) => c.borrow().item(args.first().unwrap_or(&Value::Empty)),
            Value::Object(Obj::Dictionary(d)) => Ok(d.borrow().get(args.first().unwrap_or(&Value::Empty))),
            Value::Object(Obj::Host(h)) => self.host.get_member(h, "Item", args).or_else(|_| self.host.get_member(h, "_Default", args)),
            Value::String(_) if args.is_empty() => Ok(v.clone()),
            _ => Err(err(13, "Type mismatch")),
        }
    }

    fn member(&mut self, obj: &Value, name: &str, args: &[Value], _set: bool) -> Result<Value, VbaError> {
        match obj {
            Value::Object(Obj::Host(h)) => self.host.get_member(h, name, args),
            Value::Object(Obj::Collection(c)) => self.collection_member(c, name, args),
            Value::Object(Obj::Dictionary(d)) => self.dictionary_member(d, name, args),
            Value::Object(Obj::Instance(inst)) => self.instance_member(inst, name, args),
            Value::Object(Obj::Err) => self.err_member(name),
            Value::Udt(u) => u
                .fields
                .iter()
                .find(|(f, _)| f.eq_ignore_ascii_case(name))
                .map(|(_, v)| v.clone())
                .ok_or_else(|| err(438, "Object doesn't support this property or method")),
            Value::Nothing => Err(err(91, "Object variable or With block variable not set")),
            // String members (e.g. none) or host globals offering members.
            _ => Err(err(438, format!("Object doesn't support this property or method: {name}"))),
        }
    }

    fn instance_member(&mut self, inst: &Rc<RefCell<Instance>>, name: &str, args: &[Value]) -> Result<Value, VbaError> {
        let class = inst.borrow().class.clone();
        let mi = self.program.modules.iter().position(|m| m.name.eq_ignore_ascii_case(&class));
        if let Some(mi) = mi {
            // A Property Get / method on the class.
            if let Some(pi) =
                self.program.modules.get(mi).and_then(|m| m.procs.iter().position(|p| p.name.eq_ignore_ascii_case(name) && !is_set_let(p)))
            {
                let mut slots = vec![Rc::new(RefCell::new(Value::Object(Obj::Instance(inst.clone()))))];
                slots.extend(args.iter().map(|v| Rc::new(RefCell::new(v.clone()))));
                // Bind Me via a field; simpler: run with instance fields available as globals is complex.
                return self.call_class_proc(mi, pi, inst, args);
            }
            // A public field.
            if let Some(v) = inst.borrow().fields.get(&name.to_lowercase()) {
                return Ok(v.clone());
            }
        }
        Err(err(438, format!("Object doesn't support this property or method: {name}")))
    }

    fn call_class_proc(&mut self, mi: usize, pi: usize, inst: &Rc<RefCell<Instance>>, args: &[Value]) -> Result<Value, VbaError> {
        // Fields are exposed to the procedure through a scope seeded from the instance; writes
        // flow back afterwards. (A pragmatic model of Me for single-level class code.)
        let proc = self.program.modules.get(mi).and_then(|m| m.procs.get(pi)).cloned().ok_or_else(|| err(35, "not defined"))?;
        self.depth += 1;
        if self.depth > MAX_CALL_DEPTH {
            self.depth -= 1;
            return Err(err(28, "Out of stack space"));
        }
        let slots: Vec<Slot> = args.iter().map(|v| Rc::new(RefCell::new(v.clone()))).collect();
        let r = self.run_proc_with_fields(mi, &proc, slots, inst);
        self.depth -= 1;
        r
    }

    fn run_proc_with_fields(&mut self, _mi: usize, proc: &Procedure, args: Vec<Slot>, inst: &Rc<RefCell<Instance>>) -> Result<Value, VbaError> {
        let mut scope = Scope { vars: HashMap::new(), statics: HashMap::new() };
        for (name, val) in inst.borrow().fields.iter() {
            scope.vars.insert(name.clone(), Rc::new(RefCell::new(val.clone())));
        }
        for (ai, param) in proc.params.iter().enumerate() {
            if let Some(s) = args.get(ai) {
                scope.vars.insert(param.name.to_lowercase(), if param.by_val { Rc::new(RefCell::new(s.borrow().clone())) } else { s.clone() });
            } else if let Some(e) = &param.default {
                let v = self.eval(e, &mut scope)?;
                scope.vars.insert(param.name.to_lowercase(), Rc::new(RefCell::new(v)));
            }
        }
        let returns_value = matches!(proc.kind, ProcKind::Function | ProcKind::PropertyGet);
        if returns_value {
            scope.vars.insert(proc.name.to_lowercase(), Rc::new(RefCell::new(self.type_default(proc.returns.name.as_deref()))));
        }
        let mut ctx = ProcCtx { name: proc.name.clone(), handler: None, in_handler: false };
        let flow = self.exec_lines(&proc.body, &mut scope, &mut ctx)?;
        // Write fields back.
        let field_names: Vec<String> = inst.borrow().fields.keys().cloned().collect();
        for name in field_names {
            if let Some(slot) = scope.vars.get(&name) {
                inst.borrow_mut().fields.insert(name.clone(), slot.borrow().clone());
            }
        }
        let _ = &ctx;
        match flow {
            _ if returns_value => Ok(scope.vars.get(&proc.name.to_lowercase()).map(|s| s.borrow().clone()).unwrap_or(Value::Empty)),
            _ => Ok(Value::Empty),
        }
    }

    fn collection_member(&mut self, c: &Rc<RefCell<Collection>>, name: &str, args: &[Value]) -> Result<Value, VbaError> {
        match name.to_ascii_lowercase().as_str() {
            "add" => {
                let item = args.first().cloned().unwrap_or(Value::Empty);
                let key = args.get(1).filter(|v| !matches!(v, Value::Error(_) | Value::Empty)).map(|v| v.to_str()).transpose()?;
                let before = args.get(2).filter(|v| !matches!(v, Value::Error(_) | Value::Empty)).map(|v| v.to_i64()).transpose()?;
                let after = args.get(3).filter(|v| !matches!(v, Value::Error(_) | Value::Empty)).map(|v| v.to_i64()).transpose()?;
                c.borrow_mut().add(item, key, before, after)?;
                Ok(Value::Empty)
            }
            "item" | "_default" => c.borrow().item(args.first().unwrap_or(&Value::Empty)),
            "count" => Ok(Value::from_i64(c.borrow().items.len() as i64)),
            "remove" => {
                c.borrow_mut().remove(args.first().unwrap_or(&Value::Empty))?;
                Ok(Value::Empty)
            }
            _ => Err(err(438, "Object doesn't support this property or method")),
        }
    }

    fn dictionary_member(&mut self, d: &Rc<RefCell<Dictionary>>, name: &str, args: &[Value]) -> Result<Value, VbaError> {
        match name.to_ascii_lowercase().as_str() {
            "add" => {
                let key = args.first().cloned().unwrap_or(Value::Empty);
                if d.borrow().exists(&key) {
                    return Err(err(457, "This key is already associated with an element of this collection"));
                }
                d.borrow_mut().set(key, args.get(1).cloned().unwrap_or(Value::Empty));
                Ok(Value::Empty)
            }
            "item" | "_default" => Ok(d.borrow().get(args.first().unwrap_or(&Value::Empty))),
            "exists" => Ok(Value::Bool(d.borrow().exists(args.first().unwrap_or(&Value::Empty)))),
            "count" => Ok(Value::from_i64(d.borrow().pairs.len() as i64)),
            "keys" => Ok(builtins::array_1d(d.borrow().pairs.iter().map(|(k, _)| k.clone()).collect())),
            "items" => Ok(builtins::array_1d(d.borrow().pairs.iter().map(|(_, v)| v.clone()).collect())),
            "remove" => {
                d.borrow_mut().remove(args.first().unwrap_or(&Value::Empty))?;
                Ok(Value::Empty)
            }
            "removeall" => {
                d.borrow_mut().pairs.clear();
                Ok(Value::Empty)
            }
            "comparemode" => Ok(Value::Integer(d.borrow().compare as i16)),
            _ => Err(err(438, "Object doesn't support this property or method")),
        }
    }

    fn err_member(&mut self, name: &str) -> Result<Value, VbaError> {
        let e = self.last_error.clone();
        Ok(match name.to_ascii_lowercase().as_str() {
            "number" => Value::Long(if self.err_cleared { 0 } else { e.map(|e| e.number).unwrap_or(0) }),
            "description" => Value::str(if self.err_cleared { String::new() } else { e.map(|e| e.description).unwrap_or_default() }),
            "source" => Value::str(e.map(|e| e.source).unwrap_or_default()),
            _ => Value::Empty,
        })
    }

    // ------------------------------------------------------------ assignment

    fn slot_for(&mut self, name: &str, scope: &mut Scope) -> Slot {
        let key = name.to_lowercase();
        if let Some(s) = scope.vars.get(&key) {
            return s.clone();
        }
        if let Some(s) = self.globals.get(&key) {
            return s.clone();
        }
        // Implicit variable (no Option Explicit).
        let slot = Rc::new(RefCell::new(Value::Empty));
        scope.vars.insert(key, slot.clone());
        slot
    }

    fn assign(&mut self, target: &Expr, value: Value, scope: &mut Scope, set: bool) -> Result<(), VbaError> {
        let _ = set;
        match target {
            Expr::Name(n) => {
                let slot = self.slot_for(n, scope);
                let coerced = coerce_assign(&slot.borrow(), value)?;
                *slot.borrow_mut() = coerced;
                Ok(())
            }
            Expr::Call(inner, args) => {
                // `a(i) = v`, `col(i) = v`, `obj.Prop(args) = v`, or default-member set on host.
                let argv = self.eval_args(args, scope)?;
                match &**inner {
                    Expr::Name(n) if scope.vars.contains_key(&n.to_lowercase()) || self.globals.contains_key(&n.to_lowercase()) => {
                        let slot = self.slot_for(n, scope);
                        self.assign_index(&slot, &argv, value)
                    }
                    Expr::Member(obj, m) => {
                        let o = self.eval_ref(obj, scope)?;
                        self.set_member(&o, m, &argv, value, set)
                    }
                    Expr::WithMember(m) => {
                        let o = self.with_stack.last().cloned().ok_or_else(|| err(91, "With"))?;
                        self.set_member(&o, m, &argv, value, set)
                    }
                    Expr::Name(n) => {
                        // Host default-member assignment, e.g. Cells(1,1) = v.
                        if let Some(Value::Object(Obj::Host(h))) = self.host.global(n)? {
                            return self.host.set_member(&h, "Item", &argv, value, set);
                        }
                        // Or a user procedure target — not assignable.
                        Err(err(438, format!("Can't assign to {n}(...)")))
                    }
                    _ => Err(err(13, "Type mismatch")),
                }
            }
            Expr::Member(obj, m) => {
                let o = self.eval_ref(obj, scope)?;
                self.set_member(&o, m, &[], value, set)
            }
            Expr::WithMember(m) => {
                let o = self.with_stack.last().cloned().ok_or_else(|| err(91, "With"))?;
                self.set_member(&o, m, &[], value, set)
            }
            Expr::Bang(obj, key) => {
                let o = if matches!(**obj, Expr::Name(ref n) if n == "\u{0}with") {
                    self.with_stack.last().cloned().ok_or_else(|| err(91, "With"))?
                } else {
                    self.eval_ref(obj, scope)?
                };
                self.set_member(&o, "Item", &[Value::str(key.as_str())], value, set)
            }
            _ => Err(err(13, "Type mismatch")),
        }
    }

    fn assign_index(&mut self, slot: &Slot, args: &[Value], value: Value) -> Result<(), VbaError> {
        let mut cur = slot.borrow_mut();
        match &*cur {
            Value::Array(a) => {
                let idx: Result<Vec<i64>, _> = args.iter().map(|x| x.to_i64()).collect();
                let at = a.borrow().index(&idx?)?;
                if let Some(cell) = a.borrow_mut().items.get_mut(at) {
                    *cell = value;
                }
                Ok(())
            }
            Value::Object(Obj::Dictionary(d)) => {
                d.borrow_mut().set(args.first().cloned().unwrap_or(Value::Empty), value);
                Ok(())
            }
            Value::Object(Obj::Host(h)) => {
                let h = h.clone();
                drop(cur);
                self.host.set_member(&h, "Item", args, value, false)
            }
            Value::Empty => {
                // An implicit array being filled by index isn't valid; raise.
                drop(cur);
                Err(err(9, "Subscript out of range"))
            }
            _ => {
                *cur = value;
                Ok(())
            }
        }
    }

    fn set_member(&mut self, obj: &Value, name: &str, args: &[Value], value: Value, by_ref: bool) -> Result<(), VbaError> {
        match obj {
            Value::Object(Obj::Host(h)) => self.host.set_member(h, name, args, value, by_ref),
            Value::Object(Obj::Instance(inst)) => {
                // A Property Let/Set, else a public field.
                let class = inst.borrow().class.clone();
                if let Some(mi) = self.program.modules.iter().position(|m| m.name.eq_ignore_ascii_case(&class))
                    && let Some(pi) =
                        self.program.modules.get(mi).and_then(|m| m.procs.iter().position(|p| p.name.eq_ignore_ascii_case(name) && is_set_let(p)))
                {
                    let mut a = args.to_vec();
                    a.push(value);
                    self.call_class_proc(mi, pi, inst, &a)?;
                    return Ok(());
                }
                inst.borrow_mut().fields.insert(name.to_lowercase(), value);
                Ok(())
            }
            Value::Object(Obj::Dictionary(d)) if name.eq_ignore_ascii_case("Item") => {
                d.borrow_mut().set(args.first().cloned().unwrap_or(Value::Empty), value);
                Ok(())
            }
            Value::Object(Obj::Dictionary(d)) if name.eq_ignore_ascii_case("CompareMode") => {
                d.borrow_mut().compare = value.to_i64()? as i32;
                Ok(())
            }
            Value::Udt(_) => Err(err(13, "Type mismatch")),
            Value::Nothing => Err(err(91, "Object variable or With block variable not set")),
            _ => Err(err(438, format!("Can't set {name}"))),
        }
    }

    fn eval_call_stmt(&mut self, e: &Expr, scope: &mut Scope) -> Result<(), VbaError> {
        match e {
            Expr::Call(target, args) => {
                self.eval_call(target, args, scope, true)?;
            }
            Expr::Name(n) if self.program.find_proc(n, None).is_some() && !self.local_is_indexable(n, scope) => {
                if let Some((mi, pi)) = self.program.find_proc(n, None) {
                    self.call_proc(mi, pi, vec![], true)?;
                }
            }
            Expr::Name(n) => {
                self.eval_name(n, &[], scope)?;
            }
            Expr::Member(obj, m) => {
                let o = self.eval_ref(obj, scope)?;
                self.member(&o, m, &[], false)?;
            }
            Expr::WithMember(m) => {
                let o = self.with_stack.last().cloned().ok_or_else(|| err(91, "With"))?;
                self.member(&o, m, &[], false)?;
            }
            _ => {
                self.eval(e, scope)?;
            }
        }
        Ok(())
    }
}

struct ProcCtx {
    name: String,
    handler: Option<String>,
    in_handler: bool,
}

fn is_missing(param: &Param) -> bool {
    param.optional && param.default.is_none()
}

fn ctx_static(scope: &mut Scope, _proc: &str, key: &str, make: impl FnOnce() -> Value) -> Slot {
    scope.statics.entry(key.to_string()).or_insert_with(|| Rc::new(RefCell::new(make()))).clone()
}

fn find_label(lines: &[Line], label: &str) -> Option<usize> {
    lines.iter().position(|l| matches!(l, Line::Label(n) if n == label))
}

fn lit_value(tok: &crate::lexer::Tok) -> Value {
    use crate::lexer::{NumKind, Tok};
    match tok {
        Tok::Int(n, NumKind::Integer) => Value::Integer(*n as i16),
        Tok::Int(n, NumKind::Long) => Value::Long(*n as i32),
        Tok::Int(n, NumKind::LongLong) => Value::LongLong(*n),
        Tok::Int(n, _) => Value::from_i64(*n),
        Tok::Float(f, NumKind::Single) => Value::Single(*f as f32),
        Tok::Float(f, NumKind::Currency) => Value::Currency((*f * 10_000.0).round() as i64),
        Tok::Float(f, _) => Value::Double(*f),
        Tok::Str(s) => Value::str(s.as_str()),
        Tok::Date(d) => Value::Date(*d),
        _ => Value::Empty,
    }
}

/// Coerce a value to a declared type when binding `ByVal` or assigning to a typed slot.
fn coerce(v: &Value, ty: Option<&str>) -> Result<Value, VbaError> {
    match ty.map(|t| t.to_ascii_lowercase()).as_deref() {
        Some("long") => Ok(Value::Long(v.to_i32()?)),
        Some("integer") => Ok(Value::Integer(i16::try_from(v.to_i64()?).map_err(|_| err(6, "Overflow"))?)),
        Some("byte") => Ok(Value::Byte(u8::try_from(v.to_i64()?).map_err(|_| err(6, "Overflow"))?)),
        Some("long long" | "longlong") => Ok(Value::LongLong(v.to_i64()?)),
        Some("double") => Ok(Value::Double(v.to_f64()?)),
        Some("single") => Ok(Value::Single(v.to_f64()? as f32)),
        Some("currency") => Ok(Value::Currency((v.to_f64()? * 10_000.0).round() as i64)),
        Some("boolean") => Ok(Value::Bool(v.to_bool()?)),
        Some("string") => Ok(Value::str(v.to_str()?)),
        Some("date") => Ok(Value::Date(v.to_date()?)),
        _ => Ok(v.clone()),
    }
}

/// On assignment, keep the slot's existing type when it is a fixed type (not a Variant).
fn coerce_assign(current: &Value, value: Value) -> Result<Value, VbaError> {
    let ty = match current {
        Value::Long(_) => "long",
        Value::Integer(_) => "integer",
        Value::Byte(_) => "byte",
        Value::LongLong(_) => "longlong",
        Value::Double(_) => "double",
        Value::Single(_) => "single",
        Value::Currency(_) => "currency",
        Value::Bool(_) => "boolean",
        Value::String(_) => "string",
        Value::Date(_) => "date",
        // Empty, Null, Object, Array, Udt: take the value as-is (Variant or object).
        _ => return Ok(value),
    };
    coerce(&value, Some(ty))
}

/// ReDim Preserve: copy overlapping elements from `old` into `new`.
fn preserve_array(old: Value, mut new: value::Array) -> Result<Value, VbaError> {
    if let Value::Array(a) = old {
        let a = a.borrow();
        if a.bounds.len() == new.bounds.len() {
            // Copy element-by-element over the overlapping index space.
            copy_overlap(&a, &mut new);
        }
    }
    Ok(Value::Array(Rc::new(RefCell::new(new))))
}

fn copy_overlap(old: &value::Array, new: &mut value::Array) {
    let mut idx: Vec<i64> = old.bounds.iter().map(|(lo, _)| *lo).collect();
    loop {
        if let (Ok(o), Ok(n)) = (old.index(&idx), new.index(&idx))
            && let Some(v) = old.items.get(o).cloned()
            && let Some(cell) = new.items.get_mut(n)
        {
            *cell = v;
        }
        // Increment multi-dimensional index within the overlap.
        let mut d = old.bounds.len();
        loop {
            if d == 0 {
                return;
            }
            d -= 1;
            idx[d] += 1;
            let hi = old.bounds[d].1.min(new.bounds[d].1);
            if idx[d] <= hi {
                break;
            }
            idx[d] = old.bounds[d].0;
        }
    }
}

/// The first argument, or Empty.
fn arg0(args: &[Value]) -> Value {
    args.first().cloned().unwrap_or_default()
}

/// The head of a `For var = from To to [Step step]` loop.
struct ForHead<'a> {
    var: &'a Expr,
    from: &'a Expr,
    to: &'a Expr,
    step: Option<&'a Expr>,
}

#[cfg(test)]
mod tests;
