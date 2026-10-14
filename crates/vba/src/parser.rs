//! Parsing VBA modules ([MS-VBAL] 5): declarations, procedures, statements and expressions.

use crate::ast::*;
use crate::error::{VbaError, syntax};
use crate::lexer::{Tok, Token, tokenize};
use crate::value::BinOp;

pub fn parse_module(name: &str, src: &str) -> Result<Module, VbaError> {
    let toks = tokenize(name, src)?;
    let mut p = Parser { t: toks, i: 0, module: name.to_string() };
    p.module()
}

struct Parser {
    t: Vec<Token>,
    i: usize,
    module: String,
}

fn kw(t: &Tok, k: &str) -> bool {
    matches!(t, Tok::Ident { name, bracketed: false, .. } if name.eq_ignore_ascii_case(k))
}

/// Words that can't name a variable at the start of an argument list.
const STATEMENT_ENDERS: &[&str] = &["Then", "Else", "To", "Step", "As", "In"];

impl Parser {
    fn peek(&self) -> &Tok {
        self.t.get(self.i).map(|t| &t.tok).unwrap_or(&Tok::Newline)
    }
    fn peek_at(&self, k: usize) -> &Tok {
        self.t.get(self.i + k).map(|t| &t.tok).unwrap_or(&Tok::Newline)
    }
    fn tok(&self) -> Option<&Token> {
        self.t.get(self.i)
    }
    fn line(&self) -> usize {
        self.t.get(self.i).or_else(|| self.t.last()).map_or(0, |t| t.line)
    }
    fn bump(&mut self) -> Tok {
        let t = self.peek().clone();
        self.i += 1;
        t
    }
    fn at_end(&self) -> bool {
        self.i >= self.t.len()
    }
    fn err(&self, m: impl Into<String>) -> VbaError {
        syntax(&self.module, self.line(), m)
    }
    fn is_kw(&self, k: &str) -> bool {
        kw(self.peek(), k)
    }
    fn eat_kw(&mut self, k: &str) -> bool {
        if self.is_kw(k) {
            self.i += 1;
            true
        } else {
            false
        }
    }
    fn expect_kw(&mut self, k: &str) -> Result<(), VbaError> {
        if self.eat_kw(k) { Ok(()) } else { Err(self.err(format!("Expected: {k}"))) }
    }
    fn is_op(&self, o: &str) -> bool {
        matches!(self.peek(), Tok::Op(x) if *x == o)
    }
    fn eat_op(&mut self, o: &str) -> bool {
        if self.is_op(o) {
            self.i += 1;
            true
        } else {
            false
        }
    }
    fn expect_op(&mut self, o: &str) -> Result<(), VbaError> {
        if self.eat_op(o) { Ok(()) } else { Err(self.err(format!("Expected: {o}"))) }
    }
    fn ident(&mut self) -> Result<String, VbaError> {
        match self.bump() {
            Tok::Ident { name, .. } => Ok(name),
            _ => {
                self.i -= 1;
                Err(self.err("Expected: identifier"))
            }
        }
    }
    fn end_of_statement(&self) -> bool {
        matches!(self.peek(), Tok::Newline | Tok::Colon) || self.at_end()
    }
    fn skip_separators(&mut self) {
        while matches!(self.peek(), Tok::Newline | Tok::Colon) && !self.at_end() {
            self.i += 1;
        }
    }
    fn end_statement(&mut self) -> Result<(), VbaError> {
        if self.end_of_statement() {
            if !self.at_end() {
                self.i += 1;
            }
            Ok(())
        } else {
            Err(self.err("Expected: end of statement"))
        }
    }

    // ------------------------------------------------------------ module

    fn module(&mut self) -> Result<Module, VbaError> {
        let mut m = Module { name: self.module.clone(), ..Default::default() };
        loop {
            self.skip_separators();
            if self.at_end() {
                break;
            }
            let mut public = true;
            if self.eat_kw("Option") {
                if self.eat_kw("Explicit") {
                    m.option_explicit = true;
                } else if self.eat_kw("Base") {
                    m.option_base = match self.bump() {
                        Tok::Int(n, _) => n,
                        _ => return Err(self.err("Expected: 0 or 1")),
                    };
                } else if self.eat_kw("Compare") {
                    m.compare_text = self.eat_kw("Text");
                    if !m.compare_text {
                        let _ = self.eat_kw("Binary") || self.eat_kw("Database");
                    }
                } else if self.eat_kw("Private") {
                    self.expect_kw("Module")?;
                }
                self.end_statement()?;
                continue;
            }
            if self.eat_kw("Private") {
                public = false;
            } else if self.eat_kw("Public") || self.eat_kw("Global") || self.eat_kw("Friend") {
                public = true;
            }
            let is_static = self.eat_kw("Static");
            if self.is_kw("Sub") || self.is_kw("Function") || self.is_kw("Property") {
                let proc = self.procedure(public, is_static)?;
                m.procs.push(proc);
                continue;
            }
            if self.eat_kw("Const") {
                for (n, t, e) in self.const_list()? {
                    m.consts.push((n, t, e, public));
                }
                self.end_statement()?;
                continue;
            }
            if self.eat_kw("Type") {
                let name = self.ident()?;
                self.end_statement()?;
                let mut fields = vec![];
                loop {
                    self.skip_separators();
                    if self.is_kw("End") && kw(self.peek_at(1), "Type") {
                        self.i += 2;
                        break;
                    }
                    if self.at_end() {
                        return Err(self.err("Expected: End Type"));
                    }
                    fields.push(self.var_decl()?);
                    self.end_statement()?;
                }
                m.types.push((name, fields));
                continue;
            }
            if self.eat_kw("Enum") {
                let name = self.ident()?;
                self.end_statement()?;
                let mut items = vec![];
                loop {
                    self.skip_separators();
                    if self.is_kw("End") && kw(self.peek_at(1), "Enum") {
                        self.i += 2;
                        break;
                    }
                    if self.at_end() {
                        return Err(self.err("Expected: End Enum"));
                    }
                    let n = self.ident()?;
                    let v = if self.eat_op("=") { Some(self.expr()?) } else { None };
                    items.push((n, v));
                    self.end_statement()?;
                }
                m.enums.push((name, items));
                continue;
            }
            if self.eat_kw("Declare") {
                let _ = self.eat_kw("PtrSafe");
                let _ = self.eat_kw("Sub") || self.eat_kw("Function");
                let name = self.ident()?;
                m.declares.push(name);
                while !self.end_of_statement() {
                    self.i += 1;
                }
                continue;
            }
            if self.eat_kw("Event")
                || self.eat_kw("Implements")
                || self.eat_kw("DefInt")
                || self.eat_kw("DefLng")
                || self.eat_kw("DefStr")
                || self.eat_kw("DefObj")
                || self.eat_kw("DefVar")
                || self.eat_kw("DefBool")
                || self.eat_kw("DefDbl")
            {
                while !self.end_of_statement() {
                    self.i += 1;
                }
                continue;
            }
            let _ = self.eat_kw("Dim");
            let _ = self.eat_kw("WithEvents");
            loop {
                let d = self.var_decl()?;
                m.vars.push((d, public));
                if !self.eat_op(",") {
                    break;
                }
            }
            self.end_statement()?;
        }
        Ok(m)
    }

    fn type_ref(&mut self) -> Result<TypeRef, VbaError> {
        if !self.eat_kw("As") {
            return Ok(TypeRef::default());
        }
        let new = self.eat_kw("New");
        let mut name = self.ident()?;
        while self.eat_op(".") {
            name = self.ident()?;
        }
        // `String * 10`: fixed-length strings are plain strings here.
        if self.eat_op("*") {
            let _ = self.bump();
        }
        Ok(TypeRef { name: Some(name), new })
    }

    fn bounds(&mut self) -> Result<Vec<(Option<Expr>, Expr)>, VbaError> {
        let mut out = vec![];
        if self.eat_op(")") {
            return Ok(out);
        }
        loop {
            let a = self.expr()?;
            if self.eat_kw("To") {
                let b = self.expr()?;
                out.push((Some(a), b));
            } else {
                out.push((None, a));
            }
            if !self.eat_op(",") {
                break;
            }
        }
        self.expect_op(")")?;
        Ok(out)
    }

    fn var_decl(&mut self) -> Result<VarDecl, VbaError> {
        let with_events = self.eat_kw("WithEvents");
        let name = self.ident()?;
        let dims = if self.eat_op("(") { Some(self.bounds()?) } else { None };
        let ty = self.type_ref()?;
        Ok(VarDecl { name, ty, dims, with_events })
    }

    fn const_list(&mut self) -> Result<Vec<(String, TypeRef, Expr)>, VbaError> {
        let mut out = vec![];
        loop {
            let n = self.ident()?;
            let t = self.type_ref()?;
            self.expect_op("=")?;
            out.push((n, t, self.expr()?));
            if !self.eat_op(",") {
                break;
            }
        }
        Ok(out)
    }

    fn procedure(&mut self, public: bool, is_static: bool) -> Result<Procedure, VbaError> {
        let line = self.line();
        let kind = if self.eat_kw("Sub") {
            ProcKind::Sub
        } else if self.eat_kw("Function") {
            ProcKind::Function
        } else {
            self.expect_kw("Property")?;
            if self.eat_kw("Get") {
                ProcKind::PropertyGet
            } else if self.eat_kw("Let") {
                ProcKind::PropertyLet
            } else {
                self.expect_kw("Set")?;
                ProcKind::PropertySet
            }
        };
        let name = self.ident()?;
        let mut params = vec![];
        if self.eat_op("(") && !self.eat_op(")") {
            loop {
                let mut p = Param {
                    name: String::new(),
                    by_val: false,
                    optional: false,
                    param_array: false,
                    ty: TypeRef::default(),
                    default: None,
                    is_array: false,
                };
                loop {
                    if self.eat_kw("Optional") {
                        p.optional = true;
                    } else if self.eat_kw("ByVal") {
                        p.by_val = true;
                    } else if self.eat_kw("ByRef") {
                        p.by_val = false;
                    } else if self.eat_kw("ParamArray") {
                        p.param_array = true;
                    } else {
                        break;
                    }
                }
                p.name = self.ident()?;
                if self.eat_op("(") {
                    self.expect_op(")")?;
                    p.is_array = true;
                }
                p.ty = self.type_ref()?;
                if self.eat_op("=") {
                    p.default = Some(self.expr()?);
                }
                params.push(p);
                if !self.eat_op(",") {
                    break;
                }
            }
            self.expect_op(")")?;
        }
        let returns = if matches!(kind, ProcKind::Function | ProcKind::PropertyGet) { self.type_ref()? } else { TypeRef::default() };
        self.end_statement()?;
        let end = match kind {
            ProcKind::Sub => "Sub",
            ProcKind::Function => "Function",
            _ => "Property",
        };
        let body = self.block(&[&["End", end]])?;
        self.i += 2;
        Ok(Procedure { name, kind, public, params, returns, body, line, is_static })
    }

    // ------------------------------------------------------------ statements

    /// Whether the next statement starts with one of `ends` (keyword sequences).
    fn at_terminator(&self, ends: &[&[&str]]) -> bool {
        ends.iter().any(|seq| seq.iter().enumerate().all(|(k, w)| kw(self.peek_at(k), w)))
    }

    /// Lines up to (not including) a terminator.
    fn block(&mut self, ends: &[&[&str]]) -> Result<Vec<Line>, VbaError> {
        let mut out = vec![];
        loop {
            self.skip_separators();
            if self.at_end() {
                return Err(self.err(format!("Expected: {}", ends.first().map(|s| s.join(" ")).unwrap_or_default())));
            }
            if self.at_terminator(ends) {
                return Ok(out);
            }
            self.line_into(&mut out)?;
        }
    }

    /// One statement (or label) and its separator.
    fn line_into(&mut self, out: &mut Vec<Line>) -> Result<(), VbaError> {
        // Labels: `Name:` or a line number at the start of a line.
        let line_start = self.i == 0 || matches!(self.t.get(self.i - 1).map(|t| &t.tok), Some(Tok::Newline));
        if line_start {
            if let (Tok::Ident { name, bracketed: false, .. }, Tok::Colon) = (self.peek().clone(), self.peek_at(1).clone())
                && !is_reserved(&name)
            {
                out.push(Line::Label(name.to_lowercase()));
                self.i += 2;
                return Ok(());
            }
            if let Tok::Int(n, _) = self.peek().clone() {
                out.push(Line::Label(n.to_string()));
                self.i += 1;
                let _ = matches!(self.peek(), Tok::Colon).then(|| self.i += 1);
                return Ok(());
            }
        }
        let line = self.line();
        let s = self.statement()?;
        out.push(Line::Stmt(s, line));
        // A single-line If ends its own line.
        if !self.end_of_statement() {
            return Err(self.err("Expected: end of statement"));
        }
        Ok(())
    }

    fn statement(&mut self) -> Result<Stmt, VbaError> {
        let first = self.peek().clone();
        let word = |k: &str| kw(&first, k);
        if word("Dim") || word("Static") || (word("Private") || word("Public")) && !matches!(self.peek_at(1), Tok::Op(_)) {
            let is_static = word("Static");
            self.i += 1;
            let mut decls = vec![];
            loop {
                decls.push(self.var_decl()?);
                if !self.eat_op(",") {
                    break;
                }
            }
            return Ok(Stmt::Dim { decls, is_static });
        }
        if word("Const") {
            self.i += 1;
            return Ok(Stmt::Const(self.const_list()?));
        }
        if word("ReDim") {
            self.i += 1;
            let preserve = self.eat_kw("Preserve");
            let mut items = vec![];
            loop {
                let target = {
                    let p = self.primary()?;
                    self.postfix(p, false)
                }?;
                let (target, bounds) = match target {
                    Expr::Call(t, args) => {
                        let mut b = vec![];
                        for a in args {
                            match a.value {
                                Expr::Binary(BinOp::Is, lo, hi) => b.push((Some(*lo), *hi)),
                                v => b.push((None, v)),
                            }
                        }
                        (*t, b)
                    }
                    _ => return Err(self.err("Expected: (")),
                };
                let ty = self.type_ref()?;
                items.push((target, bounds, ty));
                if !self.eat_op(",") {
                    break;
                }
            }
            return Ok(Stmt::ReDim { preserve, items });
        }
        if word("Set") {
            self.i += 1;
            let target = {
                let p = self.primary()?;
                self.postfix(p, false)
            }?;
            self.expect_op("=")?;
            return Ok(Stmt::Set { target, value: self.expr()? });
        }
        if word("Let") {
            self.i += 1;
        }
        if word("If") {
            return self.if_statement();
        }
        if word("Select") {
            self.i += 1;
            self.expect_kw("Case")?;
            let subject = self.expr()?;
            self.end_statement()?;
            let mut cases = vec![];
            let mut otherwise = None;
            loop {
                self.skip_separators();
                if self.at_terminator(&[&["End", "Select"]]) {
                    self.i += 2;
                    break;
                }
                self.expect_kw("Case")?;
                if self.eat_kw("Else") {
                    let _ = matches!(self.peek(), Tok::Colon).then(|| self.i += 1);
                    otherwise = Some(self.block(&[&["End", "Select"]])?);
                    continue;
                }
                let mut tests = vec![];
                loop {
                    if self.eat_kw("Is") {
                        let op = match self.bump() {
                            Tok::Op("=") => BinOp::Eq,
                            Tok::Op("<>") => BinOp::Ne,
                            Tok::Op("<") => BinOp::Lt,
                            Tok::Op(">") => BinOp::Gt,
                            Tok::Op("<=") => BinOp::Le,
                            Tok::Op(">=") => BinOp::Ge,
                            _ => return Err(self.err("Expected: comparison operator")),
                        };
                        tests.push(CaseTest::Is(op, self.expr()?));
                    } else {
                        let a = self.expr()?;
                        if self.eat_kw("To") {
                            tests.push(CaseTest::Range(a, self.expr()?));
                        } else {
                            tests.push(CaseTest::Value(a));
                        }
                    }
                    if !self.eat_op(",") {
                        break;
                    }
                }
                let _ = matches!(self.peek(), Tok::Colon | Tok::Newline).then(|| self.i += 1);
                let body = self.block(&[&["Case"], &["End", "Select"]])?;
                cases.push((tests, body));
            }
            return Ok(Stmt::Select { subject, cases, otherwise });
        }
        if word("For") {
            self.i += 1;
            if self.eat_kw("Each") {
                let var = {
                    let p = self.primary()?;
                    self.postfix(p, false)
                }?;
                self.expect_kw("In")?;
                let collection = self.expr()?;
                self.end_statement()?;
                let body = self.block(&[&["Next"]])?;
                self.next_clause();
                return Ok(Stmt::ForEach { var, collection, body });
            }
            let var = {
                let p = self.primary()?;
                self.postfix(p, false)
            }?;
            self.expect_op("=")?;
            let from = self.expr()?;
            self.expect_kw("To")?;
            let to = self.expr()?;
            let step = if self.eat_kw("Step") { Some(self.expr()?) } else { None };
            self.end_statement()?;
            let body = self.block(&[&["Next"]])?;
            self.next_clause();
            return Ok(Stmt::For { var, from, to, step, body });
        }
        if word("Do") {
            self.i += 1;
            let pre = self.loop_condition()?;
            self.end_statement()?;
            let body = self.block(&[&["Loop"]])?;
            self.i += 1;
            let post = self.loop_condition()?;
            return Ok(Stmt::Do { pre, post, body });
        }
        if word("While") {
            self.i += 1;
            let c = self.expr()?;
            self.end_statement()?;
            let body = self.block(&[&["Wend"]])?;
            self.i += 1;
            return Ok(Stmt::Do { pre: Some((true, c)), post: None, body });
        }
        if word("With") {
            self.i += 1;
            let object = self.expr()?;
            self.end_statement()?;
            let body = self.block(&[&["End", "With"]])?;
            self.i += 2;
            return Ok(Stmt::With { object, body });
        }
        if word("Exit") {
            self.i += 1;
            let k = match self.ident()?.to_ascii_lowercase().as_str() {
                "sub" => ExitKind::Sub,
                "function" => ExitKind::Function,
                "property" => ExitKind::Property,
                "for" => ExitKind::For,
                "do" => ExitKind::Do,
                _ => return Err(self.err("Expected: Sub, Function, Property, For or Do")),
            };
            return Ok(Stmt::Exit(k));
        }
        if word("GoTo") || word("GoSub") {
            let gosub = word("GoSub");
            self.i += 1;
            let label = self.label_name()?;
            return Ok(if gosub { Stmt::GoSub(label) } else { Stmt::GoTo(label) });
        }
        if word("Return") {
            self.i += 1;
            return Ok(Stmt::Return);
        }
        if word("On") && kw(self.peek_at(1), "Error") {
            self.i += 2;
            if self.eat_kw("Resume") {
                self.expect_kw("Next")?;
                return Ok(Stmt::OnError(OnError::ResumeNext));
            }
            self.expect_kw("GoTo")?;
            if let Tok::Int(0, _) = self.peek() {
                self.i += 1;
                return Ok(Stmt::OnError(OnError::Zero));
            }
            if matches!(self.peek(), Tok::Op("-")) {
                self.i += 2;
                return Ok(Stmt::OnError(OnError::Zero));
            }
            return Ok(Stmt::OnError(OnError::GoTo(self.label_name()?)));
        }
        if word("Resume") {
            self.i += 1;
            if self.eat_kw("Next") {
                return Ok(Stmt::Resume(ResumeKind::Next));
            }
            if self.end_of_statement() || matches!(self.peek(), Tok::Int(0, _)) {
                let _ = matches!(self.peek(), Tok::Int(..)).then(|| self.i += 1);
                return Ok(Stmt::Resume(ResumeKind::Here));
            }
            return Ok(Stmt::Resume(ResumeKind::Label(self.label_name()?)));
        }
        if word("Erase") {
            self.i += 1;
            let mut v = vec![self.expr()?];
            while self.eat_op(",") {
                v.push(self.expr()?);
            }
            return Ok(Stmt::Erase(v));
        }
        if word("End") && self.end_of_statement_at(1) {
            self.i += 1;
            return Ok(Stmt::End);
        }
        if word("Stop") {
            self.i += 1;
            return Ok(Stmt::Stop);
        }
        if word("Debug") && matches!(self.peek_at(1), Tok::Op(".")) && kw(self.peek_at(2), "Print") {
            self.i += 3;
            let mut items = vec![];
            while !self.end_of_statement() {
                if self.eat_op(";") || self.eat_op(",") {
                    continue;
                }
                items.push(self.expr()?);
            }
            return Ok(Stmt::DebugPrint(items));
        }
        if word("RaiseEvent") {
            self.i += 1;
            let n = self.ident()?;
            let args = if self.eat_op("(") { self.call_args(")")? } else { vec![] };
            return Ok(Stmt::RaiseEvent(n, args));
        }
        for w in [
            "Open",
            "Close",
            "Print",
            "Input",
            "Line",
            "Write",
            "Get",
            "Put",
            "Seek",
            "Lock",
            "Unlock",
            "Kill",
            "MkDir",
            "RmDir",
            "ChDir",
            "ChDrive",
            "Name",
            "FileCopy",
            "SendKeys",
            "Beep",
            "Randomize",
            "Load",
            "Unload",
            "AppActivate",
            "Reset",
            "SetAttr",
            "Date",
            "Time",
        ] {
            if word(w)
                && matches!(w, "Open" | "Print" | "Write" | "Get" | "Put" | "Input" | "Line" | "Seek" | "Lock" | "Unlock")
                && (matches!(self.peek_at(1), Tok::Op("#")) || w == "Open" || (w == "Line" && kw(self.peek_at(1), "Input")))
            {
                let mut text = String::new();
                while !self.end_of_statement() {
                    text.push_str(&format!("{:?} ", self.bump()));
                }
                return Ok(Stmt::Unsupported(w.to_string()));
            }
        }
        if word("Call") {
            self.i += 1;
            let e = {
                let p = self.primary()?;
                self.postfix(p, false)
            }?;
            return Ok(Stmt::Call(e));
        }
        // Assignment or call statement.
        let target = {
            let p = self.primary()?;
            self.postfix(p, true)
        }?;
        if self.eat_op("=") {
            let value = self.expr()?;
            return Ok(Stmt::Assign { target, value });
        }
        if self.end_of_statement() || self.is_kw("Else") {
            return Ok(Stmt::Call(target));
        }
        // `Foo a, b`: arguments without parentheses.
        let args = self.statement_args()?;
        Ok(Stmt::Call(Expr::Call(Box::new(target), args)))
    }

    fn end_of_statement_at(&self, k: usize) -> bool {
        matches!(self.peek_at(k), Tok::Newline | Tok::Colon) || self.i + k >= self.t.len() || kw(self.peek_at(k), "Else")
    }

    fn label_name(&mut self) -> Result<String, VbaError> {
        match self.bump() {
            Tok::Ident { name, .. } => Ok(name.to_lowercase()),
            Tok::Int(n, _) => Ok(n.to_string()),
            _ => Err(self.err("Expected: label")),
        }
    }

    fn next_clause(&mut self) {
        // `Next`, `Next i`, `Next j, i` (closing two loops is rare; the extra names are ignored).
        self.i += 1;
        while !self.end_of_statement() {
            self.i += 1;
        }
    }

    fn loop_condition(&mut self) -> Result<Option<(bool, Expr)>, VbaError> {
        if self.eat_kw("While") {
            return Ok(Some((true, self.expr()?)));
        }
        if self.eat_kw("Until") {
            return Ok(Some((false, self.expr()?)));
        }
        Ok(None)
    }

    fn if_statement(&mut self) -> Result<Stmt, VbaError> {
        self.i += 1;
        let cond = self.expr()?;
        self.expect_kw("Then")?;
        if !self.end_of_statement() {
            // Single-line If: statements up to the line end, `Else` statements after.
            let mut then = vec![];
            let mut otherwise: Option<Vec<Line>> = None;
            loop {
                let line = self.line();
                if self.eat_kw("Else") {
                    otherwise = Some(vec![]);
                    continue;
                }
                if matches!(self.peek(), Tok::Newline) || self.at_end() {
                    break;
                }
                if matches!(self.peek(), Tok::Colon) {
                    self.i += 1;
                    continue;
                }
                let s = if matches!(self.peek(), Tok::Int(..)) {
                    // `If x Then 100` is `GoTo 100`.
                    Stmt::GoTo(self.label_name()?)
                } else {
                    self.statement()?
                };
                match otherwise.as_mut() {
                    Some(o) => o.push(Line::Stmt(s, line)),
                    None => then.push(Line::Stmt(s, line)),
                }
            }
            return Ok(Stmt::If { branches: vec![(cond, then)], otherwise });
        }
        let mut branches = vec![];
        let body = self.block(&[&["ElseIf"], &["Else"], &["End", "If"]])?;
        branches.push((cond, body));
        let mut otherwise = None;
        loop {
            if self.eat_kw("ElseIf") {
                let c = self.expr()?;
                self.expect_kw("Then")?;
                let b = self.block(&[&["ElseIf"], &["Else"], &["End", "If"]])?;
                branches.push((c, b));
            } else if self.eat_kw("Else") {
                otherwise = Some(self.block(&[&["End", "If"]])?);
            } else {
                self.i += 2;
                break;
            }
        }
        Ok(Stmt::If { branches, otherwise })
    }

    /// Arguments of a call statement without parentheses: `MsgBox "x", vbOKOnly`.
    fn statement_args(&mut self) -> Result<Vec<Arg>, VbaError> {
        let mut out = vec![];
        loop {
            if self.is_op(",") {
                out.push(Arg { name: None, value: Expr::Missing, by_val: false });
                self.i += 1;
                continue;
            }
            if self.end_of_statement() || self.is_kw("Else") {
                break;
            }
            out.push(self.arg()?);
            if !self.eat_op(",") {
                break;
            }
            if self.end_of_statement() {
                out.push(Arg { name: None, value: Expr::Missing, by_val: false });
            }
        }
        Ok(out)
    }

    fn arg(&mut self) -> Result<Arg, VbaError> {
        if let (Tok::Ident { name, .. }, Tok::Op(":=")) = (self.peek().clone(), self.peek_at(1).clone()) {
            self.i += 2;
            return Ok(Arg { name: Some(name), value: self.expr()?, by_val: false });
        }
        let by_val = self.eat_kw("ByVal");
        Ok(Arg { name: None, value: self.expr()?, by_val })
    }

    /// Arguments in parentheses, up to `close`. `a To b` (ReDim bounds) becomes `Is(a, b)`.
    fn call_args(&mut self, close: &str) -> Result<Vec<Arg>, VbaError> {
        let mut out = vec![];
        if self.eat_op(close) {
            return Ok(out);
        }
        loop {
            if self.is_op(",") {
                out.push(Arg { name: None, value: Expr::Missing, by_val: false });
                self.i += 1;
                continue;
            }
            let mut a = self.arg()?;
            if self.eat_kw("To") {
                let hi = self.expr()?;
                a.value = Expr::Binary(BinOp::Is, Box::new(a.value), Box::new(hi));
            }
            out.push(a);
            if self.eat_op(",") {
                if self.is_op(close) {
                    out.push(Arg { name: None, value: Expr::Missing, by_val: false });
                }
                continue;
            }
            break;
        }
        self.expect_op(close)?;
        Ok(out)
    }

    // ------------------------------------------------------------ expressions

    pub fn expr(&mut self) -> Result<Expr, VbaError> {
        self.binary_level(0)
    }

    fn binary_level(&mut self, level: usize) -> Result<Expr, VbaError> {
        // Lowest to highest: Imp, Eqv, Xor, Or, And, Not (unary), comparisons, &, + -, Mod, \, * /, unary -, ^.
        const LEVELS: &[&[(&str, BinOp)]] =
            &[&[("imp", BinOp::Imp)], &[("eqv", BinOp::Eqv)], &[("xor", BinOp::Xor)], &[("or", BinOp::Or)], &[("and", BinOp::And)]];
        if let Some(ops) = LEVELS.get(level) {
            let mut left = self.binary_level(level + 1)?;
            while let Some(op) = ops.iter().find(|(w, _)| self.is_kw(w)).map(|x| x.1) {
                self.i += 1;
                let right = self.binary_level(level + 1)?;
                left = Expr::Binary(op, Box::new(left), Box::new(right));
            }
            return Ok(left);
        }
        self.not_level()
    }

    fn not_level(&mut self) -> Result<Expr, VbaError> {
        if self.eat_kw("Not") {
            return Ok(Expr::Unary(UnOp::Not, Box::new(self.not_level()?)));
        }
        self.comparison()
    }

    fn comparison(&mut self) -> Result<Expr, VbaError> {
        let mut left = self.concat()?;
        loop {
            let op = match self.peek() {
                Tok::Op("=") => BinOp::Eq,
                Tok::Op("<>") => BinOp::Ne,
                Tok::Op("<") => BinOp::Lt,
                Tok::Op(">") => BinOp::Gt,
                Tok::Op("<=") => BinOp::Le,
                Tok::Op(">=") => BinOp::Ge,
                t if kw(t, "Like") => BinOp::Like,
                t if kw(t, "Is") => BinOp::Is,
                _ => break,
            };
            self.i += 1;
            let right = self.concat()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn concat(&mut self) -> Result<Expr, VbaError> {
        let mut left = self.additive()?;
        while self.eat_op("&") {
            let right = self.additive()?;
            left = Expr::Binary(BinOp::Concat, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn additive(&mut self) -> Result<Expr, VbaError> {
        let mut left = self.modulo()?;
        loop {
            let op = if self.eat_op("+") {
                BinOp::Add
            } else if self.eat_op("-") {
                BinOp::Sub
            } else {
                break;
            };
            let right = self.modulo()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn modulo(&mut self) -> Result<Expr, VbaError> {
        let mut left = self.int_div()?;
        while self.eat_kw("Mod") {
            let right = self.int_div()?;
            left = Expr::Binary(BinOp::Mod, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn int_div(&mut self) -> Result<Expr, VbaError> {
        let mut left = self.multiplicative()?;
        while self.eat_op("\\") {
            let right = self.multiplicative()?;
            left = Expr::Binary(BinOp::IntDiv, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn multiplicative(&mut self) -> Result<Expr, VbaError> {
        let mut left = self.unary()?;
        loop {
            let op = if self.eat_op("*") {
                BinOp::Mul
            } else if self.eat_op("/") {
                BinOp::Div
            } else {
                break;
            };
            let right = self.unary()?;
            left = Expr::Binary(op, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn unary(&mut self) -> Result<Expr, VbaError> {
        if self.eat_op("-") {
            return Ok(Expr::Unary(UnOp::Neg, Box::new(self.unary()?)));
        }
        if self.eat_op("+") {
            return self.unary();
        }
        self.power()
    }

    fn power(&mut self) -> Result<Expr, VbaError> {
        let mut left = {
            let p = self.primary()?;
            self.postfix(p, false)
        }?;
        while self.eat_op("^") {
            // `2 ^ -1` is allowed.
            let right = if self.eat_op("-") {
                Expr::Unary(
                    UnOp::Neg,
                    Box::new({
                        let p = self.primary()?;
                        self.postfix(p, false)
                    }?),
                )
            } else {
                {
                    let p = self.primary()?;
                    self.postfix(p, false)
                }?
            };
            left = Expr::Binary(BinOp::Pow, Box::new(left), Box::new(right));
        }
        Ok(left)
    }

    fn primary(&mut self) -> Result<Expr, VbaError> {
        let spaced = self.tok().is_none_or(|t| t.spaced);
        match self.bump() {
            t @ (Tok::Int(..) | Tok::Float(..) | Tok::Str(_) | Tok::Date(_)) => Ok(Expr::Lit(t)),
            Tok::Op("(") => {
                let e = self.expr()?;
                self.expect_op(")")?;
                Ok(Expr::Call(Box::new(Expr::Name("\u{0}paren".into())), vec![Arg { name: None, value: e, by_val: true }]))
            }
            Tok::Op(".") if spaced || self.i == 1 => {
                let n = self.ident()?;
                Ok(Expr::WithMember(n))
            }
            Tok::Op("!") => {
                let n = self.ident()?;
                Ok(Expr::Bang(Box::new(Expr::Name("\u{0}with".into())), n))
            }
            Tok::Ident { name, bracketed, .. } => {
                if !bracketed {
                    if name.eq_ignore_ascii_case("New") {
                        let mut cls = self.ident()?;
                        while self.eat_op(".") {
                            cls = self.ident()?;
                        }
                        return Ok(Expr::New(cls));
                    }
                    if name.eq_ignore_ascii_case("TypeOf") {
                        let e = {
                            let p = self.primary()?;
                            self.postfix(p, false)
                        }?;
                        self.expect_kw("Is")?;
                        let mut cls = self.ident()?;
                        while self.eat_op(".") {
                            cls = self.ident()?;
                        }
                        return Ok(Expr::TypeOf(Box::new(e), cls));
                    }
                    if name.eq_ignore_ascii_case("Me") {
                        return Ok(Expr::Me);
                    }
                    if name.eq_ignore_ascii_case("AddressOf") {
                        return Err(self.err("AddressOf is not supported"));
                    }
                }
                Ok(Expr::Name(name))
            }
            _ => {
                self.i -= 1;
                Err(self.err("Expected: expression"))
            }
        }
    }

    /// Member access, `!`, and calls. In a statement head, `Foo (x)` with a space starts the
    /// arguments instead.
    fn postfix(&mut self, mut e: Expr, statement_head: bool) -> Result<Expr, VbaError> {
        loop {
            if self.is_op(".") && !self.tok().is_some_and(|t| t.spaced) {
                self.i += 1;
                let n = self.ident()?;
                e = Expr::Member(Box::new(e), n);
                continue;
            }
            if self.is_op("!") {
                self.i += 1;
                let n = self.ident()?;
                e = Expr::Bang(Box::new(e), n);
                continue;
            }
            if self.is_op("(") {
                let spaced = self.tok().is_some_and(|t| t.spaced);
                if statement_head && spaced {
                    break;
                }
                self.i += 1;
                let args = self.call_args(")")?;
                e = Expr::Call(Box::new(e), args);
                continue;
            }
            break;
        }
        let _ = STATEMENT_ENDERS;
        Ok(e)
    }
}

/// Keywords that never name a label.
fn is_reserved(name: &str) -> bool {
    ["Else", "End", "Loop", "Next", "Wend", "Case", "Resume", "Stop", "Exit", "Return", "Do"].iter().any(|k| k.eq_ignore_ascii_case(name))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn module(src: &str) -> Module {
        parse_module("Module1", src).unwrap()
    }

    #[test]
    fn procedures_and_declarations() {
        let m = module(
            "Option Explicit\r\nPrivate Const PI As Double = 3.14\r\nPublic total As Long\r\nType Pt\r\n x As Long\r\n y As Long\r\nEnd Type\r\n\r\nPublic Function Add(ByVal a As Long, Optional b As Long = 1) As Long\r\n    Add = a + b\r\nEnd Function\r\nSub Main()\r\nEnd Sub\r\n",
        );
        assert!(m.option_explicit);
        assert_eq!(m.consts.len(), 1);
        assert_eq!(m.vars[0].0.name, "total");
        assert_eq!(m.types[0].1.len(), 2);
        assert_eq!(m.procs.len(), 2);
        let f = &m.procs[0];
        assert_eq!((f.kind, f.params.len(), f.params[1].optional), (ProcKind::Function, 2, true));
    }

    #[test]
    fn statements() {
        let m = module(
            r#"Sub T()
    Dim i As Long, s As String
    For i = 1 To 10 Step 2
        If i > 5 Then Exit For Else s = s & i
    Next i
    Select Case s
        Case "a", "b": s = "x"
        Case 1 To 3
        Case Is > 4
        Case Else
            s = ""
    End Select
    With Range("A1")
        .Value = 1
        .Font.Bold = True
    End With
    Do While i > 0
        i = i - 1
    Loop
    On Error GoTo Fallo
    MsgBox "hola", vbOKOnly, "t"
    Call Foo(1, , 3)
    Range("A1:B2").Select
    Exit Sub
Fallo:
    Resume Next
End Sub"#,
        );
        let body = &m.procs[0].body;
        assert!(body.iter().any(|l| matches!(l, Line::Label(n) if n == "fallo")));
        assert!(body.iter().any(|l| matches!(l, Line::Stmt(Stmt::With { .. }, _))));
        assert!(body.iter().any(|l| matches!(l, Line::Stmt(Stmt::Select { cases, otherwise: Some(_), .. }, _) if cases.len() == 3)));
        let call = body.iter().find_map(|l| match l {
            Line::Stmt(Stmt::Call(Expr::Call(t, args)), _) if **t == Expr::Name("MsgBox".into()) => Some(args.len()),
            _ => None,
        });
        assert_eq!(call, Some(3));
    }

    #[test]
    fn precedence() {
        let mut p = Parser { t: tokenize("M", "-2 ^ 2 + 3 * 4 & \"x\" = \"a\" And Not b").unwrap(), i: 0, module: "M".into() };
        let e = p.expr().unwrap();
        let Expr::Binary(BinOp::And, l, _) = e else { panic!("{e:?}") };
        let Expr::Binary(BinOp::Eq, l, _) = *l else { panic!() };
        let Expr::Binary(BinOp::Concat, l, _) = *l else { panic!() };
        let Expr::Binary(BinOp::Add, l, _) = *l else { panic!() };
        assert!(matches!(*l, Expr::Unary(UnOp::Neg, _)));
    }

    #[test]
    fn statement_calls_with_parentheses() {
        let m = module("Sub T()\r\n    Foo (1), 2\r\n    x = Bar(1)(2)\r\n    ws.Cells(1, 2).Value = \"a\"\r\nEnd Sub");
        assert!(matches!(&m.procs[0].body[0], Line::Stmt(Stmt::Call(Expr::Call(_, a)), _) if a.len() == 2));
        assert!(matches!(&m.procs[0].body[2], Line::Stmt(Stmt::Assign { target: Expr::Member(_, n), .. }, _) if n == "Value"));
    }

    #[test]
    fn errors() {
        assert!(parse_module("M", "Sub T()\r\n If x Then\r\nEnd Sub").is_err());
        assert!(parse_module("M", "Sub T(").is_err());
        assert!(parse_module("M", "Sub T()\r\n x = \r\nEnd Sub").is_err());
    }
}
