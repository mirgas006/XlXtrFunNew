using ExcelDna.Integration;
using Python.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

public static class MyPythonFunctions
{
    private static bool pythonInitialized = false;

    // ====================================================================================
    // FINÁLNA INICIALIZAČNÁ LOGIKA PRE PRENOSITEĽNÝ ADRESÁR
    // ====================================================================================
    private static void InitializePython()
    {
        if (pythonInitialized) return;

        // 1. Získame cestu k adresáru, kde je náš .xll doplnok.
        string? xllDirectory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        if (string.IsNullOrEmpty(xllDirectory))
        {
            throw new InvalidOperationException("Could not determine the location of the XLL add-in.");
        }

        // 2. Zostavíme cestu k nášmu lokálnemu Python adresáru (vytvorenému PyInstallerom).
        string pythonRuntimePath = Path.Combine(xllDirectory, "python");
        string pythonDllPath = Path.Combine(pythonRuntimePath, "python312.dll"); // <-- DÔLEŽITÉ: Upravte na vašu verziu DLL!

        // 3. Skontrolujeme, či náš lokálny Python vôbec existuje.
        if (!Directory.Exists(pythonRuntimePath) || !File.Exists(pythonDllPath))
        {
            throw new FileNotFoundException(
                "Could not find the local Python runtime. " +
                $"Ensure the 'python' folder exists next to the .xll and contains '{Path.GetFileName(pythonDllPath)}'. " +
                $"Searched path: {pythonRuntimePath}"
            );
        }

        // 4. Nastavíme Python.NET, aby použil našu lokálnu DLL.
        Runtime.PythonDLL = pythonDllPath;

        PythonEngine.Initialize();
        using (Py.GIL())
        {
            // 5. Pridáme cestu k nášmu Python adresáru do sys.path.
            // Python si už sám nájde 'site-packages' a ostatné podadresáre.
            // Je tiež dôležité pridať aj adresár, kde je samotný .py skript.
            dynamic sys = Py.Import("sys");

            // Cesta, kde sú naše .py skripty
            string scriptDirectory = xllDirectory;

            // Pridanie ciest do Pythonu
            sys.path.append(pythonRuntimePath); // Pre nájdenie numpy, scipy atď.
            sys.path.append(scriptDirectory);
        }

        pythonInitialized = true;
    }

    // ====================================================================================
    // Pomocná funkcia na konverziu Excel Range na 1D pole
    // ====================================================================================
    private static double[] _convertRangeTo1DArray(object range)
    {
        var data = new List<double>();
        if (range is ExcelMissing) return data.ToArray();
        if (range is ExcelEmpty) return data.ToArray();
        if (range is ExcelError) return data.ToArray();

        if (range is object[,] arr2d)
        {
            foreach (var item in arr2d)
            {
                if (item is double val) data.Add(val);
            }
        }
        else if (range is double singleVal)
        {
            data.Add(singleVal);
        }
        return data.ToArray();
    }


    // ====================================================================================
    // Implementácia wrapperov pre Python funkcie (pridaná funkcia Interp)
    // ====================================================================================

    [ExcelFunction(Name = "LookupClosestValue", Description = "Returns the element in an array which is closest to a given value.")]
    public static object LookupClosestValue(
        [ExcelArgument(Name = "Array", Description = "Range of values.")] object Array,
        [ExcelArgument(Name = "ValueToSeek", Description = "The value to find the closest match for.")] double ValueToSeek)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                dynamic pythonModule = Py.Import("XlXtrFun");
                var arr = _convertRangeTo1DArray(Array);
                var result = pythonModule.LookupClosestValue(arr, ValueToSeek);
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }


    [ExcelFunction(Name = "IndexOfClosestValue", Description = "Returns the 1-based index of the element in an array which is closest to a given value.")]
    public static object IndexOfClosestValue(
        [ExcelArgument(Name = "Array", Description = "Range of values.")] object Array,
        [ExcelArgument(Name = "ValueToSeek", Description = "The value to find the closest match for.")] double ValueToSeek)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                dynamic pythonModule = Py.Import("XlXtrFun");
                var arr = _convertRangeTo1DArray(Array);
                var result = pythonModule.IndexOfClosestValue(arr, ValueToSeek);
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }

    [ExcelFunction(Name = "LookupClosestValue2D", Description = "Returns the element in a 2-D array which is closest to two given values.")]
    public static object LookupClosestValue2D(
        [ExcelArgument(Name = "XYArray", Description = "2-D range of values.")] object XYArray,
        [ExcelArgument(Name = "ArrayOfXKeys", Description = "1-D range of X keys (columns).")] object ArrayOfXKeys,
        [ExcelArgument(Name = "ArrayOfYKeys", Description = "1-D range of Y keys (rows).")] object ArrayOfYKeys,
        [ExcelArgument(Name = "XValueToSeek", Description = "The X value to seek.")] double XValueToSeek,
        [ExcelArgument(Name = "YValueToSeek", Description = "The Y value to seek.")] double YValueToSeek)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                dynamic pythonModule = Py.Import("XlXtrFun");
                var xKeys = _convertRangeTo1DArray(ArrayOfXKeys);
                var yKeys = _convertRangeTo1DArray(ArrayOfYKeys);

                // Python.NET can marshal object[,] from Excel-DNA directly to a NumPy array
                var result = pythonModule.LookupClosestValue2D(XYArray, xKeys, yKeys, XValueToSeek, YValueToSeek);
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }

    [ExcelFunction(Name = "PFit", Description = "Returns the Y at a given X on the polynomial curve of a given order least squares fit.")]
    public static object PFit(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "GivenX", Description = "The X value for which to calculate Y.")] double GivenX,
        [ExcelArgument(Name = "Order", Description = "The order of the polynomial.")] int Order,
        [ExcelArgument(Name = "Extrapolate?", Description = "Optional. TRUE to allow extrapolation.")] object Extrapolate)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                bool extrap = (Extrapolate is bool b) ? b : (Extrapolate is double d && d != 0);
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(ArrayOfXs);
                var y = _convertRangeTo1DArray(ArrayOfYs);
                var result = pythonModule.PFit(x, y, GivenX, Order, extrap);
                if (result == pythonModule.np.nan)
                {
                    return ExcelError.ExcelErrorNA;
                }
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }

    [ExcelFunction(Name = "PFitData", Description = "Returns coefficients and statistics for the polynomial curve.")]
    public static object PFitData(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "Order", Description = "The order of the polynomial.")] int Order,
        [ExcelArgument(Name = "Intercept?", Description = "Optional. FALSE forces the curve go through (0,0).")] object intercept)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                bool notRequireGoThrough00 = (intercept is bool b) ? b : (intercept is double d && d != 0);
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(ArrayOfXs);
                var y = _convertRangeTo1DArray(ArrayOfYs);

                var pyResult = pythonModule.PFitData(x, y, Order, notRequireGoThrough00);

                // --- ZAČIATOK NOVEJ LOGIKY ---

                // 1. Získame pole RIADKOV z Pythonu. Toto pole by malo mať 5 prvkov.
                var rows = pyResult.As<object[]>();

                int numRows = 5;
                int numCols = 17;

                // Kontrola, či sme dostali správny počet riadkov
                if (rows.Length != numRows)
                {
                    return $"Error: Python returned {rows.Length} rows, but expected {numRows}.";
                }

                // 2. Pripravíme si finálne 2D pole pre Excel
                var finalResult = new object[numRows, numCols];

                // 3. Prejdeme každý riadok, ktorý sme dostali z Pythonu
                for (int r = 0; r < numRows; r++)
                {
                    object rowObject = rows[r];

                    // Každý riadok je sám o sebe Python objekt, ktorý musíme premeniť na pole stĺpcov
                    if (rowObject is PyObject pyRow)
                    {
                        var cols = pyRow.As<object[]>(); // Premeníme riadok na pole stĺpcov

                        // Kontrola, či má riadok správny počet stĺpcov
                        if (cols.Length != numCols)
                        {
                            return $"Error: Row {r} has {cols.Length} columns, but expected {numCols}.";
                        }

                        // 4. Prejdeme každý stĺpec v danom riadku
                        for (int c = 0; c < numCols; c++)
                        {
                            object value = cols[c];

                            // Použijeme našu osvedčenú logiku na konverziu bunky
                            if (value == null)
                            {
                                finalResult[r, c] = ExcelEmpty.Value;
                            }
                            else if (value is PyObject pyValue)
                            {
                                finalResult[r, c] = pyValue.As<object>();
                            }
                            else
                            {
                                finalResult[r, c] = value;
                            }
                        }
                    }
                    else
                    {
                        return $"Error: Row {r} could not be converted from a Python object.";
                    }
                }

                // 5. Vrátime hotové, správne naformátované 2D pole
                return finalResult;
            }
        }
        catch (Exception ex)
        {
            return ex.ToString();
        }
    }

    [ExcelFunction(Name = "Spline", Description = "Returns the Y at a given X on the natural cubic spline curve.")]
    public static object Spline(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "GivenX", Description = "The X value(s) for which to calculate Y.")] object GivenX,
        [ExcelArgument(Name = "Extrapolate?", Description = "Optional. TRUE to allow extrapolation.")] object Extrapolate)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                bool extrap = (Extrapolate is bool b) ? b : (Extrapolate is double d && d != 0);
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(ArrayOfXs);
                var y = _convertRangeTo1DArray(ArrayOfYs);
                object givenX_py = (GivenX is double) ? GivenX : (object)_convertRangeTo1DArray(GivenX);
                var result = pythonModule.Spline(x, y, givenX_py, extrap);
                if (result == pythonModule.np.nan)
                {
                    return ExcelError.ExcelErrorNA;
                }
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }

    [ExcelFunction(Name = "Interpolate", Description = "Returns the Y at a given X on an interpolated curve.")]
    public static object Interpolate(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "GivenX", Description = "The X value(s) for which to calculate Y.")] object GivenX,
        [ExcelArgument(Name = "Extrapolate?", Description = "Optional. Default is FALSE.")] object Extrapolate,
        [ExcelArgument(Name = "Parabolic?", Description = "Optional. Default is TRUE.")] object Parabolic,
        [ExcelArgument(Name = "Averaging?", Description = "Optional. Default is TRUE.")] object Averaging,
        [ExcelArgument(Name = "SmoothingPower", Description = "Optional. Default is 1.0.")] object SmoothingPower)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                bool extrap = !(Extrapolate is ExcelMissing) && ((Extrapolate is bool b1) ? b1 : (Extrapolate is double d1 && d1 != 0));
                bool parabolic = (Parabolic is ExcelMissing) || ((Parabolic is bool b2) ? b2 : (Parabolic is double d2 && d2 != 0));
                bool averaging = (Averaging is ExcelMissing) || ((Averaging is bool b3) ? b3 : (Averaging is double d3 && d3 != 0));
                double smoothingPower = (SmoothingPower is ExcelMissing) ? 1.0 : (double)SmoothingPower;

                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(ArrayOfXs);
                var y = _convertRangeTo1DArray(ArrayOfYs);
                object givenX_py = (GivenX is double) ? GivenX : (object)_convertRangeTo1DArray(GivenX);
                var result = pythonModule.Interpolate(x, y, givenX_py, extrap, parabolic, averaging, smoothingPower);
                if (result == pythonModule.np.nan)
                {
                    return ExcelError.ExcelErrorNA;
                }
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }

    [ExcelFunction(Name = "Interp", Description = "Returns the Y at a given X on an interpolated curve using the defaults of Interpolate.")]
    public static object Interp(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "GivenX", Description = "The X value(s) for which to calculate Y.")] object GivenX)
    {
        return Interpolate(ArrayOfXs, ArrayOfYs, GivenX, false, true, true, 1.0);
    }

    [ExcelFunction(Name = "dydx", Description = "Returns the first derivative of the interpolated curve.")]
    public static object dydx(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "GivenX", Description = "The X value for which to calculate the derivative.")] double GivenX,
        [ExcelArgument(Name = "Extrapolate?", Description = "Optional. TRUE to allow extrapolation.")] object Extrapolate)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                bool extrap = (Extrapolate is bool b) ? b : (Extrapolate is double d && d != 0);
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(ArrayOfXs);
                var y = _convertRangeTo1DArray(ArrayOfYs);
                var result = pythonModule.dydx(x, y, GivenX, extrap);
                if (result == pythonModule.np.nan)
                {
                    return ExcelError.ExcelErrorNA;
                }
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }

    [ExcelFunction(Name = "ddydx", Description = "Returns the second derivative of the interpolated curve.")]
    public static object ddydx(
        [ExcelArgument(Name = "ArrayOfXs", Description = "Range of X values.")] object ArrayOfXs,
        [ExcelArgument(Name = "ArrayOfYs", Description = "Range of Y values.")] object ArrayOfYs,
        [ExcelArgument(Name = "GivenX", Description = "The X value for which to calculate the derivative.")] double GivenX,
        [ExcelArgument(Name = "Extrapolate?", Description = "Optional. TRUE to allow extrapolation.")] object Extrapolate)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                bool extrap = (Extrapolate is bool b) ? b : (Extrapolate is double d && d != 0);
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(ArrayOfXs);
                var y = _convertRangeTo1DArray(ArrayOfYs);
                var result = pythonModule.ddydx(x, y, GivenX, extrap);
                if (result == pythonModule.np.nan)
                {
                    return ExcelError.ExcelErrorNA;
                }
                return result.As<object>();
            }
        }
        catch (Exception ex) { return ex.ToString(); }
    }
    
    [ExcelFunction(Name = "XatY", Description = "Returns the X for a peak, valley, or a given Y value on the interpolated curve.")]
    public static object XatY(
        [ExcelArgument(Name = "KnownXArray", Description = "Range of X values.")] object KnownXArray,
        [ExcelArgument(Name = "KnownYArray", Description = "Range of Y values.")] object KnownYArray,
        [ExcelArgument(Name = "PeakValleyOrY", Description = "Optional. 'P' for Peak, 'V' for Valley, or 'Y' for a specific Y value. Default is 'P'.")] object PeakValleyOrY,
        [ExcelArgument(Name = "GuessX", Description = "Optional. An initial guess for the X value.")] object GuessX,
        [ExcelArgument(Name = "GivenY", Description = "Optional. The target Y value if 'PeakValleyOrY' is 'Y'.")] object GivenY,
        [ExcelArgument(Name = "Extrapolate?", Description = "Optional. TRUE to allow extrapolation.")] object Extrapolate)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                // --- Robustné spracovanie voliteľných argumentov ---

                // Pre PeakValleyOrY: Predvolená hodnota je "P".
                string solveType = "P";
                if (!(PeakValleyOrY is ExcelMissing || PeakValleyOrY is ExcelEmpty))
                {
                    // Použijeme bezpečnú konverziu. Ak by ToString() vrátilo null, použije sa "P".
                    solveType = PeakValleyOrY?.ToString()?.ToUpper() ?? "P";
                }

                // Pre GuessX a GivenY: Ak sú nevyplnené, pre Python to bude None (C# null).
                // Používame nullable typ `object?` na vyjadrenie, že hodnota môže byť null.
                object? guessX_py = null;
                if (!(GuessX is ExcelMissing || GuessX is ExcelEmpty))
                {
                    guessX_py = GuessX; // Priamo priradíme objekt, Python.NET sa postará o konverziu.
                }

                object? givenY_py = null;
                if (!(GivenY is ExcelMissing || GivenY is ExcelEmpty))
                {
                    givenY_py = GivenY;
                }

                bool extrap = (Extrapolate is bool b) ? b : (Extrapolate is double d && d != 0);

                // --- Volanie Pythonu ---
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x = _convertRangeTo1DArray(KnownXArray);
                var y = _convertRangeTo1DArray(KnownYArray);
                
                // C# `null` sa automaticky preloží na Python `None`.
                var result = pythonModule.XatY(x, y, solveType, guessX_py, givenY_py, extrap);
                return result.As<object>();
            }
        }
        catch (Exception ex)
        {
            return ex.ToString();
        }
    }

    [ExcelFunction(Name = "Intersect", Description = "Finds the intersection X-value of two curves.")]
    public static object Intersect(
        [ExcelArgument(Name = "First_Curve_Xs", Description = "X-values for the first curve.")] object First_Curve_Xs,
        [ExcelArgument(Name = "First_Curve_Ys", Description = "Y-values for the first curve.")] object First_Curve_Ys,
        [ExcelArgument(Name = "Second_Curve_Xs", Description = "X-values for the second curve.")] object Second_Curve_Xs,
        [ExcelArgument(Name = "Second_Curve_Ys", Description = "Y-values for the second curve.")] object Second_Curve_Ys,
        [ExcelArgument(Name = "Guess_X", Description = "An initial guess for the intersection's X-value.")] double Guess_X,
        [ExcelArgument(Name = "Interp_Spline_Curve_1", Description = "Optional. 'i' for Interpolate (default), 's' for Spline for the first curve.")] object Interp_Spline_Curve_1,
        [ExcelArgument(Name = "Interp_Spline_Curve_2", Description = "Optional. 'i' for Interpolate (default), 's' for Spline for the second curve.")] object Interp_Spline_Curve_2,
        [ExcelArgument(Name = "Accuracy", Description = "Optional. The desired accuracy of the result. Default is 1E-6.")] object Accuracy,
        [ExcelArgument(Name = "Max_Iterations", Description = "Optional. Maximum number of iterations. Default is 100.")] object Max_Iterations,
        [ExcelArgument(Name = "Delta_X", Description = "Optional. Small step used by the secant method. Default is 1E-3.")] object Delta_X,
        [ExcelArgument(Name = "Allow_Extrapolation", Description = "Optional. TRUE to allow extrapolation for both curves.")] object Allow_Extrapolation)
    {
        try
        {
            InitializePython();
            using (Py.GIL())
            {
                // --- Robustné spracovanie voliteľných argumentov ---
                string curve1Type = "i";
                if (!(Interp_Spline_Curve_1 is ExcelMissing || Interp_Spline_Curve_1 is ExcelEmpty))
                {
                    curve1Type = Interp_Spline_Curve_1?.ToString() ?? "i";
                }

                string curve2Type = "i";
                if (!(Interp_Spline_Curve_2 is ExcelMissing || Interp_Spline_Curve_2 is ExcelEmpty))
                {
                    curve2Type = Interp_Spline_Curve_2?.ToString() ?? "i";
                }
                
                double accuracy = 1e-6;
                if (!(Accuracy is ExcelMissing || Accuracy is ExcelEmpty))
                {
                    accuracy = Convert.ToDouble(Accuracy);
                }
                
                int maxIterations = 100;
                if (!(Max_Iterations is ExcelMissing || Max_Iterations is ExcelEmpty))
                {
                    maxIterations = Convert.ToInt32(Max_Iterations);
                }

                double deltaX = 1e-3;
                if (!(Delta_X is ExcelMissing || Delta_X is ExcelEmpty))
                {
                    deltaX = Convert.ToDouble(Delta_X);
                }

                bool allowExtrap = (Allow_Extrapolation is bool b) ? b : (Allow_Extrapolation is double d && d != 0);

                // --- Volanie Pythonu ---
                dynamic pythonModule = Py.Import("XlXtrFun");
                var x1 = _convertRangeTo1DArray(First_Curve_Xs);
                var y1 = _convertRangeTo1DArray(First_Curve_Ys);
                var x2 = _convertRangeTo1DArray(Second_Curve_Xs);
                var y2 = _convertRangeTo1DArray(Second_Curve_Ys);

                var result = pythonModule.Intersect(
                    x1, y1, x2, y2, Guess_X, 
                    curve1Type, curve2Type, 
                    accuracy, maxIterations, deltaX, allowExtrap
                );
                return result.As<object>();
            }
        }
        catch (Exception ex)
        {
            return ex.ToString();
        }
    }
}