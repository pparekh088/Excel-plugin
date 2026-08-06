/**
 * Round-trip corpus: every entry must parse with ok=true and serialize to
 * `canonical` (defaults to the input minus the leading '='). Canonical text
 * must itself re-parse to the identical serialization (fixpoint).
 */

export interface CorpusEntry {
  input: string;
  canonical?: string;
}

const e = (input: string, canonical?: string): CorpusEntry => ({ input, canonical });

export const CORPUS: CorpusEntry[] = [
  // ---- literals -------------------------------------------------------
  e("=1"),
  e("=0"),
  e("=1.5"),
  e("=.5", "0.5"),
  e("=123456789"),
  e("=3.14159"),
  e("=1E+30", "1e+30"),
  e("=2.5e-3", "0.0025"),
  e("=1e3", "1000"),
  e('="hello"'),
  e('=""'),
  e('="say ""hi"""'),
  e('="with, comma"'),
  e('="=notaformula"'),
  e("=TRUE"),
  e("=FALSE"),
  e("=true", "TRUE"),
  e("=TRUE()"),
  e("=FALSE()"),
  e("=#REF!"),
  e("=#DIV/0!"),
  e("=#N/A"),
  e("=#VALUE!"),
  e("=#NAME?"),
  e("=#NULL!"),
  e("=#NUM!"),
  e("=#SPILL!"),
  e("=#CALC!"),
  e("=#GETTING_DATA"),

  // ---- arithmetic / precedence ---------------------------------------
  e("=1+2"),
  e("=1-2"),
  e("=1*2"),
  e("=1/2"),
  e("=2^8"),
  e("=1+2*3"),
  e("=(1+2)*3"),
  e("=1*(2+3)"),
  e("=-5"),
  e("=+5"),
  e("=--5"),
  e("=-A1"),
  e("=2^-3"),
  e("=-2^2"),
  e("=2^3^2"),
  e("=1-2-3"),
  e("=1/2/2"),
  e("=2%"),
  e("=50%%"),
  e("=A1%"),
  e("=-A1%"),
  e("=1+2%"),
  e("=1=2"),
  e("=1<>2"),
  e("=1<2"),
  e("=1<=2"),
  e("=1>2"),
  e("=1>=2"),
  e("=A1>=B1"),
  e("=A1<>\"\"", '=A1<>""'.slice(1)),
  e('="a"&"b"'),
  e('=A1&B1&"x"'),
  e('="total: "&SUM(A1:A3)'),
  e("=(1)"),
  e("=((1))"),
  e("=(A1)"),
  e("=1+2+3+4+5+6+7+8+9+10"),

  // ---- cell / range refs ---------------------------------------------
  e("=A1"),
  e("=a1", "A1"),
  e("=Z9"),
  e("=AA10"),
  e("=XFD1"),
  e("=A1048576"),
  e("=$A$1"),
  e("=$A1"),
  e("=A$1"),
  e("=A1:B2"),
  e("=$A$1:$B$2"),
  e("=A1:$B$2"),
  e("=B2:A1"),
  e("=A1:A1"),
  e("=A:A"),
  e("=A:D"),
  e("=a:d", "A:D"),
  e("=$A:$B"),
  e("=$A:B"),
  e("=1:1"),
  e("=3:7"),
  e("=$1:$4"),
  e("=$2:9"),
  e("=A1:B2:C3"),
  e("=Sheet1!A1"),
  e("=Sheet1!$A$1:B2"),
  e("=Sheet1!A:A"),
  e("=Sheet1!1:3"),
  e("='My Sheet'!A1"),
  e("='It''s here'!B2"),
  e("='Sheet-1'!A1"),
  e("=Data.2024!A1"),
  e("=A1!B2", "'A1'!B2"), // sheet actually named "A1" — quoted on output
  e("='Q3'!B2"),
  e("=Sheet1:Sheet3!A1"),
  e("=Sheet1:Sheet3!A1:B10"),
  e("='Jan Data:Mar Data'!B2"),
  e("=[Book1.xlsx]Sheet1!A1", "'[Book1.xlsx]Sheet1'!A1"),
  e("='[Book1.xlsx]Sheet1'!A1"),
  e("='[Budget 2026.xlsx]Q1'!$B$2"),
  e("=[1]Sheet1!A1", "'[1]Sheet1'!A1"),
  e("=[Book1.xlsx]Sheet1:Sheet3!A1", "'[Book1.xlsx]Sheet1:Sheet3'!A1"),

  // ---- names ----------------------------------------------------------
  e("=Revenue"),
  e("=Tax_Rate"),
  e("=rev.growth"),
  e("=_hidden"),
  e("=Sheet1!LocalName"),
  e("='My Sheet'!LocalName"),
  e("=TAX*2"),
  e("=Revenue*(1+Growth_Rate)"),
  e("=TRUE_VALUE"),
  e("=XFD"), // a name, not a column (no ':')
  e("=Value?"),

  // ---- functions ------------------------------------------------------
  e("=SUM(A1:A10)"),
  e("=sum(a1:a10)", "SUM(A1:A10)"),
  e("=SUM(A1,B2,C3)"),
  e("=SUM()"),
  e("=PI()"),
  e("=NOW()"),
  e("=RAND()"),
  e('=IF(A1>0,"pos","neg")'),
  e("=IF(A1,,B1)"),
  e("=IF(A1,B1,)"),
  e("=SUM(,A1)"),
  e("=IF(A1,,)"),
  e("=IFERROR(VLOOKUP(A1,B:D,2,FALSE),0)"),
  e("=INDEX(A1:C10,2,3)"),
  e("=MATCH(A1,B:B,0)"),
  e("=INDEX(B:B,MATCH(A1,A:A,0))"),
  e("=LOG10(100)"),
  e("=ATAN2(1,2)"),
  e('=SUMIFS(C:C,A:A,">5",B:B,"x")'),
  e('=COUNTIF(A:A,"<>"&B1)'),
  e("=XLOOKUP(A1,B:B,C:C)"),
  e("=_xlfn.XLOOKUP(A1,B:B,C:C)", "XLOOKUP(A1,B:B,C:C)"),
  e("=SUMPRODUCT((A1:A10>0)*(B1:B10))"),
  e('=AI.ASK(A1,"prompt")'),
  e('=TEXT(A1,"0.00%")'),
  e('=TEXT(A1,"#,##0.00")'),
  e("=SUM(Sheet2!A1:A10,Sheet3!B:B)"),
  e("=ROUND(A1*B1,2)"),
  e("=MIN(MAX(A1,0),100)"),
  e("=NPV(0.1,B2:B10)"),
  e("=SUM(SUM(SUM(A1)))"),
  e('=CONCATENATE("a",1,TRUE)'),
  e("=N(A1)"),
  e("=T(A1)"),
  e('=IF(AND(A1>0,B1<5),OR(C1,D1),NOT(E1))'),
  e("=SUM(A1:A10)/COUNT(A1:A10)"),
  e("=VLOOKUP($A$2,Data!$A:$D,3,FALSE)"),
  e("=HLOOKUP(B1,1:4,2,TRUE)"),
  e("=SUBTOTAL(9,C2:C100)"),
  e("=SUM(A1:INDEX(A:A,10))", undefined),

  // ---- 3D and cross-sheet aggregation --------------------------------
  e("=SUM(Sheet1:Sheet3!A1)"),
  e("=SUM(Sheet1:Sheet3!A1:B10)"),
  e("=AVERAGE('Q1:Q4'!B2)"),

  // ---- modern: spill / implicit intersection / dynamic arrays --------
  e("=A1#"),
  e("=SUM(A1#)"),
  e("=Spill_Anchor#"),
  e("=@A1:A10"),
  e("=SUM(@Range1)"),
  e("=@Name"),
  e("=SEQUENCE(10)"),
  e("=FILTER(A:C,B:B>0)"),
  e("=SORT(UNIQUE(A1:A100))"),
  e("=A1#*2"),

  // ---- LET / LAMBDA ---------------------------------------------------
  e("=LET(x,1,x+1)"),
  e("=LET(x,A1,y,B1,x*y)"),
  e("=LET(x,1,y,x+1,z,y*2,z)"),
  e("=LAMBDA(x,x*2)(A1)"),
  e("=LAMBDA(a,b,a+b)(1,2)"),
  e("=BYROW(A1:C10,LAMBDA(r,SUM(r)))"),
  e("=MAP(A1:A10,LAMBDA(v,v*2))"),
  e("=LAMBDA(x,LAMBDA(y,x+y))(1)(2)"),

  // ---- array literals -------------------------------------------------
  e("={1,2;3,4}"),
  e("={1,2,3}"),
  e("={1;2;3}"),
  e('={-1,0.5;TRUE,"x"}'),
  e("={#N/A}"),
  e("=SUM({1,2,3})"),
  e("=MMULT({1,2;3,4},{5;6})"),
  e("{=SUM(A1:A3)}", "SUM(A1:A3)"),

  // ---- structured refs ------------------------------------------------
  e("=Table1[Revenue]"),
  e("=Table1[@Revenue]"),
  e("=SUM(Table1[Amount])"),
  e("=Table1[[#Headers],[Amount]]"),
  e("=Table1[[#Totals],[Amount]]"),
  e("=Table1[#All]"),
  e("=Table1[#Headers]"),
  e("=Table1[#Totals]"),
  e("=Table1[#Data]"),
  e("=Table1[Col A]"),
  e("=Table1[[Col A]]", "Table1[Col A]"),
  e("=[@Col]"),
  e("=[@[Col A]]", "[@Col A]"),
  e("=SUM(Table1[[Q1]:[Q4]])"),
  e("=Table1[@[Q1]:[Q4]]"),
  e("=SUM(Sales[Amount])/SUM(Sales[Qty])"),
  e("=T1['#Col]"),

  // ---- intersection / union ------------------------------------------
  e("=SUM(A1:A10 B5:B15)"),
  e("=B5:B15 C5:C15"),
  e("=(A1,B2)"),
  e("=SUM((A1,B2,C3))"),
  e("=COUNT((A1:A3,B1:B3))"),
  e("=A1:C10 2:2"),
  e("=(A1:A3 (B1:B3))", "(A1:A3 (B1:B3))"),
  e("=Range1 Range2"),

  // ---- computed range endpoints (opaque) ------------------------------
  e("=SUM(OFFSET(A1,1,0,10,1))"),
  e("=INDIRECT(\"A\"&B1)", '=INDIRECT("A"&B1)'.slice(1)),
  e("=INDEX(A1:C10,2,3):E5"),
  e("=SUM(A1:OFFSET(B1,5,0))"),

  // ---- whitespace tolerance ------------------------------------------
  e("= 1 + 2", "1+2"),
  e("=SUM( A1 , B2 )", "SUM(A1,B2)"),
  e("=IF( A1 > 0 , 1 , 2 )", "IF(A1>0,1,2)"),
  e("=1+\n2", "1+2"),
  e("=SUM(\n  A1:A10\n)", "SUM(A1:A10)"),
  e("=  A1  ", "A1"),

  // ---- unicode names --------------------------------------------------
  e("=Umsätze"),
  e("=收入*2"),
  e("=SUM(Ventes_Année)"),
];

// ---- generated pathological entries (still must round-trip) -----------

// 64-deep nested parens
{
  const depth = 64;
  const open = "(".repeat(depth);
  const close = ")".repeat(depth);
  CORPUS.push(e(`=${open}1${close}`));
}
// 200-term addition chain
{
  const terms = Array.from({ length: 200 }, (_, i) => String(i + 1)).join("+");
  CORPUS.push(e(`=${terms}`));
}
// SUM with 100 args
{
  const args = Array.from({ length: 100 }, (_, i) => `A${i + 1}`).join(",");
  CORPUS.push(e(`=SUM(${args})`));
}
// 32-deep IF nest
{
  let formula = "A1";
  for (let i = 0; i < 32; i++) formula = `IF(B${i + 1},${formula},C${i + 1})`;
  CORPUS.push(e(`=${formula}`));
}
// long concat chain of strings
{
  const parts = Array.from({ length: 60 }, (_, i) => `"s${i}"`).join("&");
  CORPUS.push(e(`=${parts}`));
}
