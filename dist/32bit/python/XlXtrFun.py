import numpy as np

# --- Nová trieda pre kubickú spline s použitím iba NumPy ---
class _NumpyCubicSpline:
    """
    Implementácia kubickej spline s použitím iba NumPy, inšpirovaná scipy.interpolate.CubicSpline.
    Vypočíta koeficienty pre jednotlivé kubické polynómy a umožňuje ich vyhodnocovanie.
    
    Aktuálne podporuje predovšetkým okrajovú podmienku 'natural'.
    """
    def __init__(self, x, y, bc_type='natural', extrapolate=False):
        x, y = map(np.asarray, (x, y))

        if np.issubdtype(x.dtype, np.complexfloating):
            raise ValueError("`x` must contain real values.")
        if np.iscomplexobj(y):
            # Pre jednoduchosť zatiaľ nepodporujeme komplexné y
            raise ValueError("`y` must be real for this implementation.")

        if x.ndim != 1:
            raise ValueError("`x` must be 1-dimensional.")
        
        dx = np.diff(x)
        if not np.all(dx > 0):
            # Pre zjednodušenie vyžadujeme striktne rastúce x
            raise ValueError("`x` must be strictly increasing.")

        if x.shape[0] != y.shape[0]:
            raise ValueError("The length of `y` doesn't match the length of `x`")
        
        if x.shape[0] < 2:
            raise ValueError("`x` must contain at least 2 elements.")

        self.x = x
        self.y = y
        self.dx = dx
        self.extrapolate = extrapolate
        n = len(x)

        # Pre n=2 je spline len priamka
        if n == 2:
            self.coeffs = self._solve_linear()
            return

        # Zostavenie a riešenie sústavy lineárnych rovníc pre prvé derivácie `s`
        # A * s = b
        A = np.zeros((n, n))
        b = np.zeros(n)
        
        # Vnútorné rovnice (pre i = 1, ..., n-2)
        for i in range(1, n - 1):
            A[i, i - 1] = dx[i - 1]
            A[i, i] = 2 * (dx[i - 1] + dx[i])
            A[i, i + 1] = dx[i]
            
            slope_left = (y[i] - y[i - 1]) / dx[i-1]
            slope_right = (y[i + 1] - y[i]) / dx[i]
            b[i] = 3 * (dx[i-1] * slope_right + dx[i] * slope_left)

        # Aplikácia okrajových podmienok
        if bc_type == 'natural':
            # Natural spline: druhá derivácia na koncoch je 0
            # S''_0 = 0   =>   2*s_0 + s_1 = 3 * (y_1-y_0)/h_0
            # S''_{n-1} = 0 =>   s_{n-2} + 2*s_{n-1} = 3 * (y_{n-1}-y_{n-2})/h_{n-2}
            A[0, 0] = 2
            A[0, 1] = 1
            b[0] = 3 * (y[1] - y[0]) / dx[0]

            A[n-1, n-2] = 1
            A[n-1, n-1] = 2
            b[n-1] = 3 * (y[n-1] - y[n-2]) / dx[n-2]
        else:
            # Implementácia iných okrajových podmienok by sa pridala sem
            raise NotImplementedError(f"Boundary condition '{bc_type}' is not implemented.")

        # Riešenie sústavy
        try:
            s = np.linalg.solve(A, b)
        except np.linalg.LinAlgError:
            # Fallback na pseudoinverziu, ak je matica singulárna (napr. pre rovné body)
            s = np.linalg.lstsq(A, b, rcond=None)[0]
        
        self.s = s

        # Výpočet koeficientov polynómu pre každý interval [x_i, x_{i+1}]
        # P(t) = a*t^3 + b*t^2 + c*t + d, kde t = x - x_i
        # c = s_i
        # d = y_i
        # a*h_i^3 + b*h_i^2 + c*h_i + d = y_{i+1}
        # 3a*h_i^2 + 2b*h_i + c = s_{i+1}
        
        n_intervals = n - 1
        self.coeffs = np.zeros((n_intervals, 4)) # stĺpce pre a, b, c, d
        
        for i in range(n_intervals):
            h = dx[i]
            y_i, y_i1 = y[i], y[i+1]
            s_i, s_i1 = s[i], s[i+1]

            d = y_i
            c = s_i
            # Zvyšné dva koeficienty z podmienok v bode x_{i+1}
            # a*h^3 + b*h^2 = y_i1 - s_i*h - y_i
            # 3a*h^2 + 2b*h = s_i1 - s_i
            # Riešime sústavu 2x2 pre [a, b]
            mat = np.array([[h**3, h**2], [3*h**2, 2*h]])
            rhs = np.array([y_i1 - s_i*h - y_i, s_i1 - s_i])
            try:
                a, b = np.linalg.solve(mat, rhs)
            except np.linalg.LinAlgError:
                a, b = 0, 0 # V prípade nulovej dĺžky intervalu

            self.coeffs[i] = [a, b, c, d]

    def _solve_linear(self):
        """Špeciálny prípad pre 2 body (lineárna interpolácia)."""
        coeffs = np.zeros((1, 4))
        slope = (self.y[1] - self.y[0]) / self.dx[0]
        coeffs[0, 2] = slope # c (lineárny koeficient)
        coeffs[0, 3] = self.y[0] # d (konštanta)
        return coeffs

    def __call__(self, given_x):
        """Vyhodnotí spline v bode/bodoch given_x."""
        is_scalar = np.isscalar(given_x)
        given_x = np.atleast_1d(given_x)
        
        # Príprava výstupného poľa
        results = np.full(given_x.shape, np.nan)

        if not self.extrapolate:
            in_bounds = (given_x >= self.x[0]) & (given_x <= self.x[-1])
            x_eval = given_x[in_bounds]
        else:
            in_bounds = np.ones_like(given_x, dtype=bool)
            x_eval = given_x

        if x_eval.size == 0:
            return results[0] if is_scalar else results

        # Nájdenie indexov intervalov pre všetky body naraz
        # `i` bude index ľavého bodu intervalu [x_i, x_{i+1}]
        indices = np.searchsorted(self.x, x_eval, side='right') - 1
        
        # Ošetrenie krajných bodov
        indices[indices < 0] = 0
        indices[indices >= len(self.coeffs)] = len(self.coeffs) - 1

        # Získanie koeficientov pre príslušné intervaly
        active_coeffs = self.coeffs[indices]
        # Výpočet lokálnej súradnice t = x - x_i
        t = x_eval - self.x[indices]

        # Vyhodnotenie polynómov vektorizovaným spôsobom
        # a*t^3 + b*t^2 + c*t + d
        y_vals = active_coeffs[:, 0] * t**3 + \
                 active_coeffs[:, 1] * t**2 + \
                 active_coeffs[:, 2] * t + \
                 active_coeffs[:, 3]
        
        results[in_bounds] = y_vals
        
        return results[0] if is_scalar else results


# --- Pomocné funkcie ---
def _get_interval_index(x_array, given_x):
    """Nájde index i, pre ktorý platí x_array[i] <= given_x < x_array[i+1]"""
    if given_x < x_array[0]:
        return -1
    if given_x > x_array[-1]:
        return len(x_array) - 1
    
    idx = np.searchsorted(x_array, given_x, side='right')
    if idx == 0: return 0
    if idx < len(x_array) and x_array[idx] == given_x:
        return idx
    return idx - 1

# --- Hlavné funkcie ---
def LookupClosestValue(Array, ValueToSeek):
    """Vráti prvok v poli, ktorý je najbližšie k danej hodnote."""
    arr = np.asarray(Array)
    idx = np.argmin(np.abs(arr - ValueToSeek))
    return arr[idx]

def IndexOfClosestValue(Array, ValueToSeek):
    """Vráti 1-založený index prvku v poli, ktorý je najbližšie k danej hodnote."""
    arr = np.asarray(Array)
    idx = np.argmin(np.abs(arr - ValueToSeek))
    return idx + 1

def LookupClosestValue2D(XYArray, ArrayOfXKeys, ArrayOfYKeys, XValueToSeek, YValueToSeek):
    """Vráti prvok v 2D poli na základe najbližších kľúčov X a Y."""
    xy_arr = np.asarray(XYArray)
    x_keys = np.asarray(ArrayOfXKeys)
    y_keys = np.asarray(ArrayOfYKeys)
    
    x_idx = np.argmin(np.abs(x_keys - XValueToSeek))
    y_idx = np.argmin(np.abs(y_keys - YValueToSeek))
    
    return xy_arr[y_idx, x_idx]

def PFit(ArrayOfXs, ArrayOfYs, GivenX, Order, Extrapolate=False):
    """
    Vráti Y pre daný X na polynomiálnej krivke.
    Používa numpy.linalg.lstsq pre výpočet metódou najmenších štvorcov.
    """
    x = np.array(ArrayOfXs)
    y = np.array(ArrayOfYs)
    
    if not Extrapolate and (GivenX < x.min() or GivenX > x.max()):
        return np.nan

    # Vytvorenie matice s mocninami X (X^1, X^2, ..., X^Order)
    X_poly = np.vander(x, Order + 1, increasing=True)[:, 1:]
    
    # Pridanie konštanty (stĺpec jednotiek) pre výpočet interceptu
    X_design = np.c_[np.ones(x.shape[0]), X_poly]
        
    # Vytvorenie a fitovanie modelu pomocou NumPy
    # np.linalg.lstsq rieši rovnicu X @ c = y pre c
    coeffs, _, _, _ = np.linalg.lstsq(X_design, y, rcond=None)
    
    # Pre výpočet hodnoty musíme koeficienty otočiť, aby zodpovedali formátu pre np.polyval
    # np.polyval očakáva [c_n, c_{n-1}, ..., c_1, c_0]
    # lstsq vracia [c_0, c_1, ..., c_n]
    coeffs_for_polyval = np.flip(coeffs)
    
    return np.polyval(coeffs_for_polyval, GivenX)

def PFitData(ArrayOfXs, ArrayOfYs, Order, RequireGoThrough00=False):
    """
    Vráti koeficienty a štatistiky pre polynomiálnu krivku, podobne ako LINEST v Exceli.
    Implementované pomocou NumPy, bez závislosti na statsmodels.

    Args:
        ArrayOfXs (list or np.array): Pole X súradníc.
        ArrayOfYs (list or np.array): Pole Y súradníc.
        Order (int): Rád polynómu.
        RequireGoThrough00 (bool): Ekvivalent `const` v LINEST. Ak False, krivka prechádza (0,0).

    Returns:
        np.array: 5x17 pole obsahujúce koeficienty a štatistiky.
    """
    x = np.array(ArrayOfXs)
    y = np.array(ArrayOfYs)
    n = len(x)
    
    # Vytvorenie matice s mocninami X
    # np.vander s increasing=True vytvorí stĺpce [x^0, x^1, ..., x^Order]
    # Ponecháme si stĺpce od x^1, stĺpec pre x^0 (konštantu) pridáme podľa potreby
    X_vander = np.vander(x, Order + 1, increasing=True)

    if not RequireGoThrough00:
        # Štandardný model s interceptom
        # Dizajnová matica X má stĺpce [1, x, x^2, ..., x^Order]
        X_design = X_vander
        p = Order + 1 # Počet parametrov (koeficientov) vrátane interceptu
    else:
        # Model nútený prejsť bodom (0,0) - bez interceptu
        # Dizajnová matica X má stĺpce [x, x^2, ..., x^Order]
        X_design = X_vander[:, 1:]
        p = Order # Počet parametrov
        
    # Výpočet koeficientov pomocou metódy najmenších štvorcov
    coeffs, residuals, rank, s = np.linalg.lstsq(X_design, y, rcond=None)
    
    # --- Výpočet štatistík ---
    
    # Suma štvorcov rezíduí (ss_resid)
    # residuals je pole, zoberieme prvý prvok
    ss_resid = residuals[0] if residuals.size > 0 else 0.0

    # Stupne voľnosti rezíduí (df_resid)
    df_resid = n - p
    
    # Priemerná kvadratická chyba (Mean Squared Error)
    mse_resid = ss_resid / df_resid if df_resid > 0 else 0.0
    
    # Štandardná chyba odhadu Y (se_y)
    se_y = np.sqrt(mse_resid)

    # Kovaiančná matica koeficientov: mse * (X'X)^-1
    try:
        xtx_inv = np.linalg.inv(X_design.T @ X_design)
        cov_matrix = mse_resid * xtx_inv
        # Štandardné chyby koeficientov sú odmocniny diagonálnych prvkov
        std_errs = np.sqrt(np.diag(cov_matrix))
    except np.linalg.LinAlgError:
        # Matica je singulárna, nemôžeme vypočítať chyby
        std_errs = np.full(p, np.nan)

    # Celková suma štvorcov (ss_total)
    if not RequireGoThrough00:
        # Pre model s interceptom sa počíta odchýlka od priemeru
        ss_total = np.sum((y - y.mean())**2)
    else:
        # Pre model bez interceptu sa počíta odchýlka od nuly
        ss_total = np.sum(y**2)

    # Regresná suma štvorcov (ss_reg)
    ss_reg = ss_total - ss_resid
    
    # Koeficient determinácie (R^2)
    rsquared = ss_reg / ss_total if ss_total > 0 else 0.0

    # F-štatistika
    df_reg = p - 1 if not RequireGoThrough00 else p
    ms_reg = ss_reg / df_reg if df_reg > 0 else 0.0
    f_value = ms_reg / mse_resid if mse_resid > 0 else 0.0
    
    # --- Zostavenie výstupného poľa ---
    output = np.zeros((5, 17))
    
    # Riadok 1 & 2: Koeficienty a ich štandardné chyby
    if not RequireGoThrough00:
        #coeffs sú [c0, c1, c2, ...], čo je správne poradie
        output[0, :p] = coeffs
        output[1, :p] = std_errs
    else:
        # Pre model bez interceptu je c0 = 0.
        # coeffs sú [c1, c2, ...], vložíme ich od druhého stĺpca
        output[0, 1:p+1] = coeffs
        output[1, 1:p+1] = std_errs
        
    # Riadok 3: r^2 a std. chyba y
    output[2, 0] = rsquared
    output[2, 1] = se_y
    
    # Riadok 4: F-štatistika a stupne voľnosti
    output[3, 0] = f_value
    output[3, 1] = df_resid
    
    # Riadok 5: Suma štvorcov (regresná a reziduálna)
    output[4, 0] = ss_reg
    output[4, 1] = ss_resid
    
    return output


def Spline(ArrayOfXs, ArrayOfYs, GivenX, Extrapolate=False):
    """Vráti Y pre daný X na prirodzenej kubickej spline krivke."""
    x = np.array(ArrayOfXs)
    y = np.array(ArrayOfYs)

    # Skontrolujeme, či je pole monotónne (rastúce alebo klesajúce)
    dx = np.diff(x)
    is_increasing = np.all(dx >= 0)
    is_decreasing = np.all(dx <= 0)

    if not is_increasing and not is_decreasing:
        raise ValueError("ArrayOfXs must be monotonic (constantly increasing or decreasing).")

    # Ak je pole klesajúce, otočíme ho (spolu s y), aby sme mohli použiť rovnaký algoritmus
    if is_decreasing:
        x = x[::-1]
        y = y[::-1]
    
    # Vytvorenie a vyhodnotenie spline pomocou našej NumPy implementácie
    cs = _NumpyCubicSpline(x, y, bc_type='natural', extrapolate=Extrapolate)
    result = cs(GivenX)

    return result

def Interpolate(ArrayOfXs, ArrayOfYs, GivenX, Extrapolate=False, Parabolic=True, Averaging=True, SmoothingPower=1.0):
    """Vráti Y pre daný X na interpolovanej krivke."""
    x = np.array(ArrayOfXs)
    y = np.array(ArrayOfYs)
    n = len(x)

    if n < 2: raise ValueError("Interpolation requires at least two points.")
    # Skontrolujeme, či je pole monotónne (rastúce alebo klesajúce)
    dx = np.diff(x)
    is_increasing = np.all(dx >= 0)
    is_decreasing = np.all(dx <= 0)

    if not is_increasing and not is_decreasing:
        raise ValueError("ArrayOfXs must be monotonic (constantly increasing or decreasing).")

    # Ak je pole klesajúce, otočíme ho (spolu s y), aby sme mohli použiť rovnaký algoritmus
    if is_decreasing:
        x = x[::-1]
        y = y[::-1]

    if not Parabolic:
        if not Extrapolate:
            return np.interp(GivenX, x, y, left=np.nan, right=np.nan)
        else:
            # Rozdelenie logiky pre skalár a pole
            if np.isscalar(GivenX):
                # Logika pre JEDNU hodnotu
                if GivenX < x[0]:
                    slope_left = (y[1] - y[0]) / (x[1] - x[0])
                    return y[0] + slope_left * (GivenX - x[0])
                elif GivenX > x[-1]:
                    slope_right = (y[-1] - y[-2]) / (x[-1] - x[-2])
                    return y[-1] + slope_right * (GivenX - x[-1])
                else:
                    return np.interp(GivenX, x, y)
            else:
                # Logika pre POLE hodnôt
                given_x_np = np.asarray(GivenX)
                result = np.interp(given_x_np, x, y) # Interpoluje, kraje "upne"

                # Identifikácia bodov na extrapoláciu
                is_extrap_left = given_x_np < x[0]
                is_extrap_right = given_x_np > x[-1]

                # Aplikácia extrapolácie na príslušné časti poľa
                if np.any(is_extrap_left):
                    slope_left = (y[1] - y[0]) / (x[1] - x[0])
                    result[is_extrap_left] = y[0] + slope_left * (given_x_np[is_extrap_left] - x[0])
                if np.any(is_extrap_right):
                    slope_right = (y[-1] - y[-2]) / (x[-1] - x[-2])
                    result[is_extrap_right] = y[-1] + slope_right * (given_x_np[is_extrap_right] - x[-1])
                return result

    # Parabolická interpolácia
    if n < 3:
        return np.interp(GivenX, x, y, left=np.nan, right=np.nan) if not Extrapolate else np.interp(GivenX, x, y)

    i = _get_interval_index(x, GivenX)
    if np.isscalar(GivenX) and GivenX == x[-1]: i = n - 2

    if i == -1 or i >= n - 1:
        if not Extrapolate: return np.nan
        points_x, points_y = (x[:3], y[:3]) if i == -1 else (x[-3:], y[-3:])
        poly = np.polyfit(points_x, points_y, 2)
        return np.polyval(poly, GivenX)

    x_left, y_left = (x[0:3], y[0:3]) if i == 0 else (x[i-1:i+2], y[i-1:i+2])
    x_right, y_right = (x[n-3:n], y[n-3:n]) if i >= n - 2 else (x[i:i+3], y[i:i+3])
    
    poly_left = np.polyfit(x_left, y_left, 2)
    poly_right = np.polyfit(x_right, y_right, 2)

    if not Averaging:
        return np.polyval(poly_left, GivenX)
    
    Y_left_parabola, Y_right_parabola = np.polyval(poly_left, GivenX), np.polyval(poly_right, GivenX)
    knot_left_x, knot_right_x = x[i], x[i+1]
    if knot_right_x == knot_left_x: return y[i]
        
    weight = ((GivenX - knot_left_x) / (knot_right_x - knot_left_x)) ** SmoothingPower
    return Y_right_parabola * weight + Y_left_parabola * (1 - weight)

def dydx(ArrayOfXs, ArrayOfYs, GivenX, Extrapolate=False):
    """Vráti prvú deriváciu interpolovanej krivky."""
    x = np.array(ArrayOfXs)
    if (np.min(GivenX) < x.min() or np.max(GivenX) > x.max()) and not Extrapolate:
        return np.nan # Zjednodušená kontrola pre pole
    """w, dw, y_l, y_r, dy_l, dy_r, _, _ = _get_derivative_components(ArrayOfXs, ArrayOfYs, GivenX)
    return dy_r * w + y_r * dw + dy_l * (1 - w) - y_l * dw"""
    h = 1e-6
    f1 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX - h, True)
    f2 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX + h, True)
    f3 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX - 2 * h, True)
    f4 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX + 2 * h, True)
    der = (-f4 + 8 * f2 - 8 * f1 + f3) / (12 * h)
    return der

def ddydx(ArrayOfXs, ArrayOfYs, GivenX, Extrapolate=False):
    """Vráti druhú deriváciu interpolovanej krivky."""
    x = np.array(ArrayOfXs)
    if (np.min(GivenX) < x.min() or np.max(GivenX) > x.max()) and not Extrapolate:
        return np.nan # Zjednodušená kontrola pre pole
    """w, dw, _, _, dy_l, dy_r, ddy_l, ddy_r = _get_derivative_components(ArrayOfXs, ArrayOfYs, GivenX)
    return ddy_r * w + ddy_l * (1 - w) + 2 * (dy_r - dy_l) * dw"""
    h = 1e-4
    f0 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX, True)
    f1 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX - h, True)
    f2 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX + h, True)
    f3 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX - 2 * h, True)
    f4 = Interpolate(ArrayOfXs, ArrayOfYs, GivenX + 2 * h, True)
    der = (-f4 + 16 * f2 - 30 * f0 + 16 * f1 - f3) / (12 * h * h)
    return der

def XatY(
    KnownXArray,
    KnownYArray,
    PeakValleyOrY='P',
    GuessX=None,
    GivenY=None,
    Extrapolate=False,
):
    """
    Vráti X pre maximum (Peak), minimum (Valley) alebo dané Y.
    - Pre Peak/Valley hľadá koreň prvej derivácie pomocou metódy sečníc.
    - Pre dané Y hľadá koreň funkcie (Interpolate(x) - GivenY) pomocou metódy sečníc.
    """
    accuracy=1e-5,
    max_iterations=100
    solve_type = PeakValleyOrY.upper()
    known_xs = np.array(KnownXArray)
    known_ys = np.array(KnownYArray)
    
    # Definovanie funkcie pre metódu sečníc
    if solve_type in ['P', 'V']:
        target_func = lambda x: dydx(known_xs, known_ys, x, Extrapolate)
    elif solve_type == 'Y':
        if GivenY is None:
            raise ValueError("Pre 'PeakValleyOrY=Y' je potrebné zadať 'GivenY'.")
        target_func = lambda x: Interpolate(known_xs, known_ys, x, Extrapolate) - GivenY
    else:
        raise ValueError("Neplatný 'PeakValleyOrY'. Použite 'P', 'V' alebo 'Y'.")

    # Nastavenie počiatočného odhadu (`GuessX`)
    if GuessX is None:
        if solve_type == 'P':
            GuessX = known_xs[np.nanargmax(known_ys)]
        elif solve_type == 'V':
            GuessX = known_xs[np.nanargmin(known_ys)]
        elif solve_type == 'Y':
            GuessX = known_xs[np.nanargmin(np.abs(known_ys - GivenY))]

    # Implementácia metódy sečníc
    x0 = GuessX
    # Malý posun pre druhý bod
    delta_x = (np.max(known_xs) - np.min(known_xs)) / 1000 or 1e-4
    x1 = x0 + delta_x

    f0 = target_func(x0)
    f1 = target_func(x1)
    
    if np.isnan(f0) or np.isnan(f1):
        raise RuntimeError(f"Nepodarilo sa vypočítať funkciu v okolí GuessX={GuessX:.4f}. Skúste iný odhad.")

    for _ in range(max_iterations):
        if abs(f1 - f0) < 1e-14:
            # Ak sú hodnoty príliš blízko, konvergencia sa zastavila
            return x1
        
        # Vzorec metódy sečníc
        x_next = x1 - f1 * (x1 - x0) / (f1 - f0)
        
        if abs(x_next - x1) < accuracy:
            # Overenie typu extrému, ak bol hľadaný
            if solve_type in ['P', 'V']:
                second_deriv = ddydx(known_xs, known_ys, x_next, Extrapolate)
                is_peak = second_deriv < 0
                is_valley = second_deriv > 0
                
                if solve_type == 'P' and not is_peak:
                    raise RuntimeError(f"Nájdený extrém v X={x_next:.4f} nie je maximum (ddydx={second_deriv:.4f}).")
                if solve_type == 'V' and not is_valley:
                    raise RuntimeError(f"Nájdený extrém v X={x_next:.4f} nie je minimum (ddydx={second_deriv:.4f}).")
            
            return x_next
        
        x0, f0 = x1, f1
        x1, f1 = x_next, target_func(x_next)
        
        if np.isnan(f1):
            raise RuntimeError(f"Iterácia metódy sečníc viedla na neplatnú hodnotu (NaN). Skúste iný GuessX.")

    raise RuntimeError(f"Metóda nekonvergovala po {max_iterations} iteráciách.")

def Intersect(
    First_Curve_Xs,
    First_Curve_Ys,
    Second_Curve_Xs,
    Second_Curve_Ys,
    Guess_X,
    Interp_Spline_Curve_1='i',
    Interp_Spline_Curve_2='i',
    Accuracy=1e-5,
    Max_Iterations=100,
    Delta_X=1e-3,
    Allow_Extrapolation=False
):
    """
    Nájde priesečník dvoch kriviek pomocou metódy sečníc.
    """
    curve1_type = Interp_Spline_Curve_1.lower()
    curve2_type = Interp_Spline_Curve_2.lower()
    func1 = Spline if 's' in curve1_type else Interpolate
    func2 = Spline if 's' in curve2_type else Interpolate

    def difference_func(x):
        y1 = func1(First_Curve_Xs, First_Curve_Ys, x, Extrapolate=Allow_Extrapolation)
        y2 = func2(Second_Curve_Xs, Second_Curve_Ys, x, Extrapolate=Allow_Extrapolation)
        return y1 - y2

    x0 = Guess_X
    x1 = Guess_X + Delta_X
    
    f0 = difference_func(x0)
    f1 = difference_func(x1)

    for _ in range(Max_Iterations):
        if abs(f1 - f0) < 1e-15: return x1
        x_next = x1 - f1 * (x1 - x0) / (f1 - f0)
        if abs(x_next - x1) < Accuracy:
            return x_next
        x0, f0 = x1, f1
        x1, f1 = x_next, difference_func(x_next)
        
    raise RuntimeError(f"Metóda sečníc pre intersect nekonvergovala po {Max_Iterations} iteráciách.")