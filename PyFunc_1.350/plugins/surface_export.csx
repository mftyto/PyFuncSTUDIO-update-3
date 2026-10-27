// C# extension: export a 3D table to Excel (.xlsx) or CSV.
// The user picks which two coordinates form the grid; the selected formula
// computes the remaining coordinate. The formula can use a, b, t and the
// usual x, y, z names.

using PyFunc.PluginApi;

string pair = ctx.Settings.Choice("pair", "XY", "Координаты (a, b)", "XY", "XZ", "YZ");
string formula = ctx.Settings.Text("formula", "sin(a) * cos(b)", "Формула третьей координаты");
double aFrom = ctx.Settings.Number("aFrom", -3.0, "a: начало", -1e6, 1e6, 0.1, "%.4g");
double aTo = ctx.Settings.Number("aTo", 3.0, "a: конец", -1e6, 1e6, 0.1, "%.4g");
double aStep = ctx.Settings.Number("aStep", 0.5, "a: шаг", 0.0001, 1e5, 0.05, "%.4g");
double bFrom = ctx.Settings.Number("bFrom", -3.0, "b: начало", -1e6, 1e6, 0.1, "%.4g");
double bTo = ctx.Settings.Number("bTo", 3.0, "b: конец", -1e6, 1e6, 0.1, "%.4g");
double bStep = ctx.Settings.Number("bStep", 0.5, "b: шаг", 0.0001, 1e5, 0.05, "%.4g");
double precisionValue = ctx.Settings.Number("precision", 3, "Знаков после запятой", 0, 12, 1, "%.0f");
string format = ctx.Settings.Choice("format", "xlsx", "Формат файла", "xlsx", "csv");
string path = ctx.Settings.File("path", TableExport.DefaultPath("pyfunc_surface_table.xlsx"), "Файл");
ctx.Settings.Button("export", "Сформировать 3D-таблицу");

if (ctx.WasAction("export"))
{
    try
    {
        int precision = (int)Math.Round(Math.Clamp(precisionValue, 0.0, 12.0));
        if (aStep <= 0.0 || bStep <= 0.0 || aTo < aFrom || bTo < bFrom)
            throw new InvalidOperationException("Проверьте интервалы и шаги a/b");

        pair = pair.ToUpperInvariant();
        string c1, c2, c3;
        switch (pair)
        {
            case "XZ": c1 = "x"; c2 = "z"; c3 = "y"; break;
            case "YZ": c1 = "y"; c2 = "z"; c3 = "x"; break;
            default: pair = "XY"; c1 = "x"; c2 = "y"; c3 = "z"; break;
        }

        var expression = new MathExpression(formula);
        var rows = new List<double?[]>();
        int guard = 0;
        for (double a = aFrom; a <= aTo + aStep * 1e-9 && guard < 250_000; a += aStep, guard++)
        {
            int inRow = 0;
            for (double b = bFrom; b <= bTo + bStep * 1e-9 && guard < 250_000; b += bStep, guard++, inRow++)
            {
                double x = 0.0, y = 0.0, z = 0.0;
                if (pair == "XY") { x = a; y = b; }
                else if (pair == "XZ") { x = a; z = b; }
                else { y = a; z = b; }

                double value = expression.Evaluate(x, y, z, ctx.Time, a, b);
                double? third = double.IsFinite(value) ? value : null;
                double?[] row = pair switch
                {
                    "XZ" => new double?[] { x, z, third },
                    "YZ" => new double?[] { y, z, third },
                    _ => new double?[] { x, y, third },
                };
                rows.Add(row);
            }
            if (inRow == 0) break;
        }

        if (rows.Count == 0) throw new InvalidOperationException("Нет точек для таблицы");
        string saved = TableExport.Save(path, format, new[] { c1, c2, c3 }, rows, precision);
        ctx.Notify($"3D-таблица {pair} сохранена: {saved} ({rows.Count} строк)");
        ctx.Log("C# 3D-таблица: " + saved);
    }
    catch (Exception ex)
    {
        ctx.Notify("Ошибка 3D-таблицы: " + ex.Message);
        ctx.Log("C# 3D-таблица: ошибка — " + ex.Message);
    }
}
