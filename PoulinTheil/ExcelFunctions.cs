using ExcelDna.Integration;
using System.Collections.Generic;
using System.Linq;
using static System.Math;

namespace PoulinTheil
{
  public class ExcelFunctions
  {
    [ExcelFunction(Description = "Pt:p at steady state, assuming a homogenous distribution and passive diffusion of drugs in tissues, from a ratio of solubility and macromolecular binding between tissues and plasma.")]
    public static object Ptp_Eq11(
      [ExcelArgument("rat|rabbit|mouse")] string species,
      [ExcelArgument("brain|heart|lung|muscle|skin|intestine|spleen|bone")] string tissue,
      [ExcelArgument("Log vegetable oil:water partition coefficient")] double logKvow,
      [ExcelArgument("Unbound fraction in lipid non-depleted plasma")] double fup
      )
    {
      species = species.Trim().ToLowerInvariant();
      tissue = tissue.Trim().ToLowerInvariant();

      if (tissue == "plasma") return ExcelError.ExcelErrorValue;
      if(!_Vw.ContainsKey(tissue) || !_Vw.First().Value.ContainsKey(species)) return ExcelError.ExcelErrorValue;
      var Vwt = _Vw[tissue][species];
      var Vnt = _Vn[tissue][species];
      var Vpht = _Vph[tissue][species];
      if (double.IsNaN(Vwt) || double.IsNaN(Vnt) || double.IsNaN(Vpht)) return ExcelError.ExcelErrorNA;

      var Vwp = _Vw["plasma"][species];
      var Vnp = _Vn["plasma"][species];
      var Vphp = _Vph["plasma"][species];

      if (fup <= 0.0) return ExcelError.ExcelErrorDiv0;
      var nKaCmp = (1.0 / fup) - 1.0;
      var nKaCmt = nKaCmp * 0.5;
      var fut = 1.0 / (1.0 + nKaCmt);

      var Kvow = Pow(10, logKvow);

      var StSw = Kvow * (Vnt + 0.3 * Vpht) + (Vwt + 0.7 * Vpht);
      var SpSw = Kvow * (Vnp + 0.3 * Vphp) + (Vwp + 0.7 * Vphp);

      var P = (StSw / SpSw) * (fup / fut);

      return P;
    }

    [ExcelFunction(Description = "Pt:p, assuming non-homogenous distribution, of drugs residing predominantly in the interstitial space of tissues.")]
    public static object Ptp_Eq14(
      [ExcelArgument("brain|heart|lung|muscle|skin|intestine|spleen|bone")] string tissue,
      [ExcelArgument("Unbound fraction in lipid non-depleted plasma")] double fup
      )
    {
      tissue = tissue.Trim().ToLowerInvariant();

      if (tissue == "plasma") return ExcelError.ExcelErrorValue;
      if (!_Ft.ContainsKey(tissue)) return ExcelError.ExcelErrorValue;
      var Ft = _Ft[tissue];

      if (fup <= 0.0) return ExcelError.ExcelErrorDiv0;
      var nKaCmp = (1.0 / fup) - 1.0;
      var nKaCmt = nKaCmp * 0.5;
      var fut = 1.0 / (1.0 + nKaCmt);

      var P = (Ft / _Fp) * (fup / fut);

      return P;
    }

    private static IDictionary<string, IDictionary<string, double>> _Vw = new SortedDictionary<string, IDictionary<string, double>>
    {
      ["plasma"] = new SortedDictionary<string, double> { ["rabbit"] = 0.94, ["rat"] = 0.96, ["mouse"] = 0.96 },
      ["brain"] = new SortedDictionary<string, double> { ["rabbit"] = 0.74, ["rat"] = 0.75, ["mouse"] = 0.71 },
      ["heart"] = new SortedDictionary<string, double> { ["rabbit"] = 0.79, ["rat"] = 0.77, ["mouse"] = 0.78 },
      ["lung"] = new SortedDictionary<string, double> { ["rabbit"] = 0.78, ["rat"] = 0.78, ["mouse"] = 0.81 },
      ["muscle"] = new SortedDictionary<string, double> { ["rabbit"] = 0.77, ["rat"] = 0.74, ["mouse"] = 0.67 },
      ["skin"] = new SortedDictionary<string, double> { ["rabbit"] = 0.70, ["rat"] = 0.70, ["mouse"] = double.NaN },
      ["intestine"] = new SortedDictionary<string, double> { ["rabbit"] = 0.70, ["rat"] = 0.70, ["mouse"] = 0.70 },
      ["spleen"] = new SortedDictionary<string, double> { ["rabbit"] = double.NaN, ["rat"] = 0.77, ["mouse"] = 0.79 },
      ["bone"] = new SortedDictionary<string, double> { ["rabbit"] = 0.35, ["rat"] = 0.35, ["mouse"] = double.NaN },
    };

    private static IDictionary<string, IDictionary<string, double>> _Vn = new SortedDictionary<string, IDictionary<string, double>>
    {
      ["plasma"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0015, ["rat"] = 0.00147, ["mouse"] = 0.0026 },
      ["brain"] = new SortedDictionary<string, double> { ["rabbit"] = 0.05, ["rat"] = 0.0393, ["mouse"] = 0.031 },
      ["heart"] = new SortedDictionary<string, double> { ["rabbit"] = 0.031, ["rat"] = 0.0117, ["mouse"] = 0.017 },
      ["lung"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0332, ["rat"] = 0.0199, ["mouse"] = 0.0218 },
      ["muscle"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0338, ["rat"] = 0.009, ["mouse"] = 0.0167 },
      ["skin"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0205, ["rat"] = 0.0205, ["mouse"] = double.NaN },
      ["intestine"] = new SortedDictionary<string, double> { ["rabbit"] = 0.032, ["rat"] = 0.032, ["mouse"] = 0.032 },
      ["spleen"] = new SortedDictionary<string, double> { ["rabbit"] = double.NaN, ["rat"] = 0.0077, ["mouse"] = 0.0120 },
      ["bone"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0222, ["rat"] = 0.0222, ["mouse"] = double.NaN },
    };

    private static IDictionary<string, IDictionary<string, double>> _Vph = new SortedDictionary<string, IDictionary<string, double>>
    {
      ["plasma"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0012, ["rat"] = 0.00083, ["mouse"] = 0.0032 },
      ["brain"] = new SortedDictionary<string, double> { ["rabbit"] = 0.064, ["rat"] = 0.0532, ["mouse"] = 0.05 },
      ["heart"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0082, ["rat"] = 0.0141, ["mouse"] = 0.014 },
      ["lung"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0142, ["rat"] = 0.0170, ["mouse"] = 0.0162 },
      ["muscle"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0062, ["rat"] = 0.01, ["mouse"] = 0.0273 },
      ["skin"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0155, ["rat"] = 0.0155, ["mouse"] = double.NaN },
      ["intestine"] = new SortedDictionary<string, double> { ["rabbit"] = 0.015, ["rat"] = 0.015, ["mouse"] = 0.015 },
      ["spleen"] = new SortedDictionary<string, double> { ["rabbit"] = double.NaN, ["rat"] = 0.0136, ["mouse"] = 0.0107 },
      ["bone"] = new SortedDictionary<string, double> { ["rabbit"] = 0.0005, ["rat"] = 0.0005, ["mouse"] = double.NaN },
    };

    private static IDictionary<string, double> _Ft = new SortedDictionary<string, double>
    {
      ["brain"] = 0.17,
      ["heart"] = 0.1,
      ["lung"] = 0.188,
      ["muscle"] = 0.12,
      ["skin"] = 0.302,
      ["intestine"] = 0.094,
      ["spleen"] = 0.15,
      ["bone"] = 0.1,
    };

    private static double _Fp = 0.63;
  }
}
