// ============================================================================
// Power Query (M) voor de cashflow-template
// Power BI Desktop > Gegevens transformeren > Nieuwe bron > Lege query >
// Geavanceerde editor. Maak twee queries: "Cashflow" en "Instellingen".
// Pas BronPad aan naar de locatie van Cashflow_Template.xlsx
// (of maak er een parameter van via Parameters beheren).
// ============================================================================

// ---------- Query: Cashflow ----------
let
    BronPad   = "C:\Pad\Naar\Cashflow_Template.xlsx",
    Bron      = Excel.Workbook(File.Contents(BronPad), null, true),
    Tabel     = Bron{[Item = "tblCashflow", Kind = "Table"]}[Data],
    Typen     = Table.TransformColumnTypes(Tabel, {
                    {"Kwartaal", type date}, {"Versie", type text},
                    {"Opbrengsten", type number}, {"Kosten", type number}}),
    Leeg      = Table.SelectRows(Typen, each [Kwartaal] <> null and [Versie] <> null),
    // Elke datum naar de eerste dag van het kwartaal
    Kwartaal  = Table.TransformColumns(Leeg, {{"Kwartaal", Date.StartOfQuarter, type date}}),
    Versie    = Table.TransformColumns(Kwartaal, {{"Versie", Text.Trim, type text}}),
    // Ontbrekende bedragen = 0, kosten altijd negatief
    Nullen    = Table.ReplaceValue(Versie, null, 0, Replacer.ReplaceValue, {"Opbrengsten", "Kosten"}),
    Kosten    = Table.TransformColumns(Nullen, {{"Kosten", each -Number.Abs(_), type number}}),
    // Meerdere regels per kwartaal/versie (bijv. per project) optellen
    Groep     = Table.Group(Kosten, {"Kwartaal", "Versie"}, {
                    {"Opbrengsten", each List.Sum([Opbrengsten]), type number},
                    {"Kosten",      each List.Sum([Kosten]),      type number}})
in
    Groep

// ---------- Query: Instellingen ----------
let
    BronPad   = "C:\Pad\Naar\Cashflow_Template.xlsx",
    Bron      = Excel.Workbook(File.Contents(BronPad), null, true),
    Tabel     = Bron{[Item = "tblInstellingen", Kind = "Table"]}[Data],
    Draai     = Table.PromoteHeaders(Table.Transpose(Tabel), [PromoteAllScalars = true]),
    Typen     = Table.TransformColumnTypes(Draai, {{"RealisatieTotEnMet", type date}}),
    Kwartaal  = Table.TransformColumns(Typen, {{"RealisatieTotEnMet", Date.StartOfQuarter, type date}})
in
    Kwartaal
