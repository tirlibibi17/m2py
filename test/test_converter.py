
import re
from m2py_core import convert_m_to_python

def norm(s): return re.sub(r'\s+', ' ', s.strip())

def test_expand_record():
    m = \"\"\"
    let
        Source = Table.FromRecords({[A=1, Rec=[x=10,y=20]], [A=2, Rec=[x=30,y=40]]}),
        Expanded = Table.ExpandRecordColumn(Source, "Rec", {"x","y"}, {"X","Y"})
    in
        Expanded
    \"\"\"
    py = convert_m_to_python(m, query_name="Q")
    assert "ExpandRecordColumn" in py or "join(_exp)" in py

def test_expand_table():
    m = \"\"\"
    let
        T1 = #table(type table [x=number, y=number], {{10,20}}),
        Source = #table(type table [A=number, Tbl=table], {{1, T1}}),
        Expanded = Table.ExpandTableColumn(Source, "Tbl", {"x","y"}, {"X","Y"})
    in
        Expanded
    \"\"\"
    py = convert_m_to_python(m, query_name="Q")
    assert "ExpandTableColumn" in py or "explode()" in py

def test_csv_options():
    m = \"\"\"
    let
        Source = Csv.Document(File.Contents("data.csv"), [Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
        Promoted = Table.PromoteHeaders(Source)
    in
        Promoted
    \"\"\"
    py = convert_m_to_python(m, "Q")
    assert "pd.read_csv('data.csv', header=None" in py and "sep=';'" in py and "encoding='utf-8'" in py

def test_multikey_join():
    m = \"\"\"
    let
        A = Excel.CurrentWorkbook(){[Name="A"]}[Content],
        B = Excel.CurrentWorkbook(){[Name="B"]}[Content],
        J = Table.Join(A, {"k1","k2"}, B, {"k1","k2"}, "Bcols", JoinKind.Inner)
    in
        J
    \"\"\"
    py = convert_m_to_python(m, "Q")
    assert "merge(" in py and "left_on=['k1', 'k2']" in py

def test_type_mapping():
    m = \"\"\"
    let
        Source = Excel.CurrentWorkbook(){[Name="T"]}[Content],
        Types = Table.TransformColumnTypes(Source, {{"txt", type text}, {"num", type number}, {"dt", type datetime}, {"flag", type logical}})
    in
        Types
    \"\"\"
    py = convert_m_to_python(m, "Q")
    assert "astype(" in py or "_TYPE_MAP" in py
