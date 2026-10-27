// C# extension: configurable table export to Excel (.xlsx) or CSV.
// All controls below are rendered in the "C# РАСШИРЕНИЯ" panel.
// The table is built only when the button is pressed.

using PyFunc.PluginApi;

string formula = ctx.Settings.Text("formula", "sin(x) + 0.35*cos(3*x)", "Формула f(x)");
double from = ctx.Settings.Number("from", -6.0, "Начало интервала", -1e6, 1e6, 0.1, "%.4g");
double to = ctx.Settings.Number("to", 6.0, "Конец интервала", -1e6, 1e6, 0.1, "%.4g");
double step = ctx.Settings.Number("step", 0.25, "Шаг", 0.0001, 1e5, 0.05, "%.4g");
string individual = ctx.Settings.Text("values", "0; 1.5; 3.14159", "Отдельные x (через ;)");
double precisionValue = ctx.Settings.Number("precision", 3, "Знаков после запятой", 0, 12, 1, "%.0f");
string format = ctx.Settings.Choice("format", "xlsx", "Формат файла", "xlsx", "csv");
string path = ctx.Settings.File("path", TableExport.DefaultPath("pyfunc_table.xlsx"), "Файл");
ctx.Settings.Button("export", "Сформировать таблицу");

if (ctx.WasAction("export"))
{
    try
    {
        int precision = (int)Math.Round(Math.Clamp(precisionValue, 0.0, 12.0));
        var expression = new MathExpression(formula);
        var xs = new SortedSet<double>();

        if (step > 0.0 && to >= from)
        {
            int guard = 0;
            for (double x = from; x <= to + step * 1e-9 && guard < 200_000; x += step, guard++)
            {
                xs.Add(x);
            }
        }

        foreach (double x in ParseValues(individual))
        {
            if (xs.Count >= 200_000) break;
            xs.Add(x);
        }

        if (xs.Count == 0) throw new InvalidOperationException("Нет точек для таблицы");

        var rows = new List<double?[]>(xs.Count);
        foreach (double x in xs)
        {
            double y = expression.Evaluate(x, t: ctx.Time);
            rows.Add(new double?[] { x, double.IsFinite(y) ? y : null });
        }

        string saved = TableExport.Save(path, format, new[] { "x", "f(x)" }, rows, precision);
        ctx.Notify($"Таблица сохранена: {saved} ({rows.Count} строк)");
        ctx.Log("C# таблица: " + saved);
    }
    catch (Exception ex)
    {
        ctx.Notify("Ошибка таблицы: " + ex.Message);
        ctx.Log("C# таблица: ошибка — " + ex.Message);
    }
}

static IEnumerable<double> ParseValues(string text)
{
    foreach (string raw in text.Split(new[] { ';', '\n', '\r', '\t', ' ' },
                                      StringSplitOptions.RemoveEmptyEntries))
    {
        string token = raw.Trim().Replace(',', '.');
        if (double.TryParse(token, out double value) && double.IsFinite(value))
        {
            yield return value;
        }
    }
}
