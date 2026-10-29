//! The syntax tree of a VBA module ([MS-VBAL] 5).

use crate::value::BinOp;

#[derive(Clone, Debug, PartialEq)]
pub enum Expr {
    Lit(crate::lexer::Tok),
    /// A bare name (variable, constant, procedure without arguments, global object).
    Name(String),
    /// `.Member` inside a `With` block.
    WithMember(String),
    /// `object.Member`.
    Member(Box<Expr>, String),
    /// `object!Key` (the default member with a string argument).
    Bang(Box<Expr>, String),
    /// `target(args)`: a call, an array index or a default member.
    Call(Box<Expr>, Vec<Arg>),
    Unary(UnOp, Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
    New(String),
    /// `TypeOf x Is Class`.
    TypeOf(Box<Expr>, String),
    /// A missing optional argument (`f(1, , 3)`).
    Missing,
    Me,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum UnOp {
    Neg,
    Not,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Arg {
    /// `name:=value`.
    pub name: Option<String>,
    pub value: Expr,
    /// `ByVal x` at a call site.
    pub by_val: bool,
}

/// A declared type (`As Long`, `As New Collection`, `As String * 10`).
#[derive(Clone, Debug, PartialEq, Default)]
pub struct TypeRef {
    pub name: Option<String>,
    pub new: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct VarDecl {
    pub name: String,
    pub ty: TypeRef,
    /// `Dim a(1 To 3, 5)`: bounds; `Dim a()`: dynamic (`Some(vec![])`).
    pub dims: Option<Vec<(Option<Expr>, Expr)>>,
    pub with_events: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub enum CaseTest {
    Value(Expr),
    Range(Expr, Expr),
    Is(BinOp, Expr),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    Property,
    For,
    Do,
}

#[derive(Clone, Debug, PartialEq)]
pub enum OnError {
    GoTo(String),
    ResumeNext,
    Zero,
}

#[derive(Clone, Debug, PartialEq)]
pub enum ResumeKind {
    Here,
    Next,
    Label(String),
}

#[derive(Clone, Debug, PartialEq)]
pub enum Stmt {
    Dim {
        decls: Vec<VarDecl>,
        is_static: bool,
    },
    Const(Vec<(String, TypeRef, Expr)>),
    ReDim {
        preserve: bool,
        items: Vec<ReDimItem>,
    },
    /// `x = 1`, `Range("A1").Value = 1`, `Mid(s, 2, 1) = "x"`.
    Assign {
        target: Expr,
        value: Expr,
    },
    /// `Set x = obj`.
    Set {
        target: Expr,
        value: Expr,
    },
    /// A call statement: `Foo 1, 2`, `Call Foo(1, 2)`, `obj.Method x`.
    Call(Expr),
    If {
        branches: Vec<(Expr, Vec<Line>)>,
        otherwise: Option<Vec<Line>>,
    },
    Select {
        subject: Expr,
        cases: Vec<(Vec<CaseTest>, Vec<Line>)>,
        otherwise: Option<Vec<Line>>,
    },
    For {
        var: Expr,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Line>,
    },
    ForEach {
        var: Expr,
        collection: Expr,
        body: Vec<Line>,
    },
    /// `Do [While|Until c] … Loop [While|Until c]`; `While … Wend` too.
    Do {
        pre: Option<(bool, Expr)>,
        post: Option<(bool, Expr)>,
        body: Vec<Line>,
    },
    With {
        object: Expr,
        body: Vec<Line>,
    },
    Exit(ExitKind),
    GoTo(String),
    GoSub(String),
    Return,
    OnError(OnError),
    Resume(ResumeKind),
    Erase(Vec<Expr>),
    /// `Debug.Print a; b`.
    DebugPrint(Vec<Expr>),
    End,
    Stop,
    RaiseEvent(String, Vec<Arg>),
    /// `Open … For … As #n`, `Print #n`, `Close`: not supported yet; reported when run.
    Unsupported(String),
}

/// A statement with its source line, or a label (`Retry:`, `100`).
#[derive(Clone, Debug, PartialEq)]
pub enum Line {
    Stmt(Stmt, usize),
    Label(String),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Param {
    pub name: String,
    pub by_val: bool,
    pub optional: bool,
    pub param_array: bool,
    pub ty: TypeRef,
    pub default: Option<Expr>,
    pub is_array: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub struct Procedure {
    pub name: String,
    pub kind: ProcKind,
    pub public: bool,
    pub params: Vec<Param>,
    pub returns: TypeRef,
    pub body: Vec<Line>,
    pub line: usize,
    pub is_static: bool,
}

#[derive(Clone, Debug, PartialEq, Default)]
pub struct Module {
    pub name: String,
    pub option_explicit: bool,
    pub option_base: i64,
    pub compare_text: bool,
    /// Module-level variables (`Dim`, `Private`, `Public`, `Global`).
    pub vars: Vec<(VarDecl, bool)>,
    pub consts: Vec<(String, TypeRef, Expr, bool)>,
    pub types: Vec<(String, Vec<VarDecl>)>,
    pub enums: Vec<EnumDef>,
    pub procs: Vec<Procedure>,
    /// `Declare` statements (Windows API calls): kept by name to report them.
    pub declares: Vec<String>,
}

/// One array of a `ReDim`: the variable, its bounds (`lower To upper`, lower optional) and type.
pub type ReDimItem = (Expr, Vec<(Option<Expr>, Expr)>, TypeRef);

/// An `Enum` block: its name and members, each with an optional value.
pub type EnumDef = (String, Vec<(String, Option<Expr>)>);
