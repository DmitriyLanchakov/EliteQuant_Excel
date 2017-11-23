using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Xl = Microsoft.Office.Interop.Excel;
using EliteQuant;

namespace EliteQuantExcel.Models
{
    public class Math
    {
        #region Threefry
        [ExcelFunction(Description = "Generate Threefry PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenThreefryURng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            Xl.Range range = ExcelUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            try
            {
                double[,] ret = new double[c,1];
                double[] ret0 = new double[c];

                EliteQuant.NQuantLibc.uniformthreefry(seed, skip, ret0, c);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = ret0[i];
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Threefry PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenThreefryGaussianRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            Xl.Range range = ExcelUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            try
            {
                double[,] ret = new double[c, 1];
                double[] ret0 = new double[c];

                EliteQuant.NQuantLibc.normalthreefry(seed, skip, ret0, c);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = ret0[i];
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Threefry PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenThreefryGammaRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip,
            [ExcelArgument(Description = "shape (shape) ")]double shape,
                [ExcelArgument(Description = "scale (scale) ")]double scale)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();
            Xl.Range range = ExcelUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            try
            {
                double[,] ret = new double[c, 1];
                double[] ret0 = new double[c];

                EliteQuant.NQuantLibc.gammathreefry(seed, skip, ret0, c, shape, scale);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = ret0[i];
                }

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }
        /*
        [ExcelFunction(Description = "Generate Threefry PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenThreefryGaussianRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostThreefryUniformRng rng = new BoostThreefryUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);
                EliteQuant.InvCumulativeThreefryGaussianRng rngN = new InvCumulativeThreefryGaussianRng(rng);
                
                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rngN.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Threefry PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenThreefryStudentRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip,
            [ExcelArgument(Description = "Student t Param ")]int N)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostThreefryUniformRng rng = new BoostThreefryUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);
                InverseCumulativeStudent ic = new InverseCumulativeStudent(N);
                EliteQuant.InvCumulativeThreefryStudentRng rngT = new InvCumulativeThreefryStudentRng(rng, ic);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rngT.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Threefry PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenThreefryPoissonRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip,
            [ExcelArgument(Description = "Poisson Param ")]double lambda)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostThreefryUniformRng rng = new BoostThreefryUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);
                InverseCumulativePoisson ic = new InverseCumulativePoisson(lambda);
                EliteQuant.InvCumulativeThreefryPoissonRng rngP = new InvCumulativeThreefryPoissonRng(rng, ic);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rngP.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }
        #endregion

        #region Philox
        [ExcelFunction(Description = "Generate Philox PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenPhiloxURng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostPhiloxUniformRng rng = new BoostPhiloxUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rng.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Philox PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenPhiloxGaussianRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostPhiloxUniformRng rng = new BoostPhiloxUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);
                EliteQuant.InvCumulativePhiloxGaussianRng rngN = new InvCumulativePhiloxGaussianRng(rng);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rngN.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Philox PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenPhiloxStudentRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip,
            [ExcelArgument(Description = "Student t Param ")]int N)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostPhiloxUniformRng rng = new BoostPhiloxUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);
                InverseCumulativeStudent ic = new InverseCumulativeStudent(N);
                EliteQuant.InvCumulativePhiloxStudentRng rngT = new InvCumulativePhiloxStudentRng(rng, ic);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rngT.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Generate Philox PRNG", Category = "EliteQuantExcel - Math")]
        public static object eqMathGenPhiloxPoissonRng(
            [ExcelArgument(Description = "seed of the rng (can't be zero)")] int seed,
            [ExcelArgument(Description = "(restart)counterbase ")]int counterbase,
            [ExcelArgument(Description = "skip (jump forward) ")]int skip,
            [ExcelArgument(Description = "Poisson Param ")]double lambda)
        {
            if (SystemUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = SystemUtil.getActiveCellAddress();
            Xl.Range range = SystemUtil.getActiveCellRange();
            int c = range.Count;        // range should be one column

            if (seed == 0)
            {
                Random r = new Random();
                counterbase = r.Next(10000);
            }

            try
            {
                double[,] ret = new double[c, 1];
                EliteQuant.BoostPhiloxUniformRng rng = new BoostPhiloxUniformRng(seed);
                rng.restart(counterbase);
                rng.discard(skip);
                InverseCumulativePoisson ic = new InverseCumulativePoisson(lambda);
                EliteQuant.InvCumulativePhiloxPoissonRng rngP = new InvCumulativePhiloxPoissonRng(rng, ic);

                for (int i = 0; i < c; i++)
                {
                    ret[i, 0] = rngP.next().value();
                }

                return ret;
            }
            catch (Exception e)
            {
                SystemUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }*/
        #endregion

        #region Interpolation
        [ExcelFunction(Description = "One Dimensional interpolation", Category = "EliteQuantExcel - Math")]
        public static string eqMathLinearInterpolation(
            [ExcelArgument(Description = "interpolation obj id")] string objId,
            [ExcelArgument(Description = "variable x ")]double[] x,
            [ExcelArgument(Description = "variable y ")]double[] y)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (x.Length != y.Length)
                {
                    return "size mismatch";
                }

                QlArray xa = new QlArray((uint)x.Length);
                QlArray ya = new QlArray((uint)y.Length);

                for (uint i = 0; i < x.Length; i++)
                {
                    xa.set(i, x[i]);
                    ya.set(i, y[i]);
                }

                LinearInterpolation interp = new LinearInterpolation(xa, ya);

                // Store the futures and return its id
                string id = "Int@" + objId;
                OHRepository.Instance.storeObject(id, interp, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "One Dimensional interpolation", Category = "EliteQuantExcel - Math")]
        public static string eqMathLogLinearInterpolation(
            [ExcelArgument(Description = "interpolation obj id")] string objId,
            [ExcelArgument(Description = "variable x ")]double[] x,
            [ExcelArgument(Description = "variable y ")]double[] y)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (x.Length != y.Length)
                {
                    return "size mismatch";
                }

                QlArray xa = new QlArray((uint)x.Length);
                QlArray ya = new QlArray((uint)y.Length);

                for (uint i = 0; i < x.Length; i++)
                {
                    xa.set(i, x[i]);
                    ya.set(i, y[i]);
                }

                LogLinearInterpolation interp = new LogLinearInterpolation(xa, ya);

                // Store the futures and return its id
                string id = "Int@" + objId;
                OHRepository.Instance.storeObject(id, interp, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "One Dimensional interpolation", Category = "EliteQuantExcel - Math")]
        public static object eqMathGet1DInterpolation(
            [ExcelArgument(Description = "interpolation obj id")] string objId,
            [ExcelArgument(Description = "variable x ")]double x,
            [ExcelArgument(Description = "Linear/LogLinear ")]string type)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (type.ToUpper() == "LINEAR")
                {
                    LinearInterpolation interp = OHRepository.Instance.getObject<LinearInterpolation>(objId);
                    double ret = interp.call(x, true);

                    return ret;
                }
                else if (type.ToUpper() == "LOGLINEAR")
                {
                    LogLinearInterpolation interp = OHRepository.Instance.getObject<LogLinearInterpolation>(objId);
                    double ret = interp.call(x, true);

                    return ret;
                }
                else
                {
                    return "Unknown interpolation type";
                }
                
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Two Dimensional interpolation", Category = "EliteQuantExcel - Math")]
        public static string eqMathBiLinearInterpolation(
            [ExcelArgument(Description = "interpolation obj id")] string objId,
            [ExcelArgument(Description = "row variable x ")]double[] x,
            [ExcelArgument(Description = "col variable y ")]double[] y,
            [ExcelArgument(Description = "variable z ")]double[,] z)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if ((x.Length != z.GetLength(0)) || (y.Length != z.GetLength(1)))
                {
                    return "size mismatch";
                }

                QlArray xa = new QlArray((uint)x.Length);
                QlArray ya = new QlArray((uint)y.Length);
                for (uint i = 0; i < x.Length; i++)
                {
                    xa.set(i, x[i]);
                    ya.set(i, y[i]);
                }

                Matrix ma = new Matrix((uint)x.Length, (uint)y.Length);
                for (uint i = 0; i < x.Length; i++)
                {
                    for (uint j = 0; j < y.Length; j++)
                    {
                        ma.set(i, j, z[i, j]);
                    }
                }

                BilinearInterpolation interp = new BilinearInterpolation(xa, ya, ma);

                // Store the futures and return its id
                string id = "Int@" + objId;
                OHRepository.Instance.storeObject(id, interp, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Two Dimensional interpolation", Category = "EliteQuantExcel - Math")]
        public static object eqMathGet2DInterpolation(
            [ExcelArgument(Description = "interpolation obj id")] string objId,
            [ExcelArgument(Description = "row variable x ")]double x,
            [ExcelArgument(Description = "col variable y ")]double y)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                BilinearInterpolation interp = OHRepository.Instance.getObject<BilinearInterpolation>(objId);

                double ret = interp.call(x, y);
                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }
        #endregion
    }
}
