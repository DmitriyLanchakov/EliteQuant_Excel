using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;
using Xl = Microsoft.Office.Interop.Excel;
using EliteQuant;

namespace EliteQuantExcel
{
    public class Volatility
    {
        #region SABR
        [ExcelFunction(Description = "SABR section calibration", Category = "EliteQuantExcel - Models")]
        public static string eqModelSABR(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "ATM forward ")]double forward,
            [ExcelArgument(Description = "expiry (years of fraction ")]double tenor,
            [ExcelArgument(Description = "Strikes ")]double[] strikes,
            [ExcelArgument(Description = "implied vols ")]double[] volatilities,
            [ExcelArgument(Description = "initial parameters ")]double[] initials,
            [ExcelArgument(Description = "is parameters fixed? ")]object[] isfixed,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (strikes.Length != volatilities.Length)
                    throw new Exception("SABR input lengths don't match.");

                if ((initials.Length !=4) || (isfixed.Length != 4))
                    throw new Exception("SABR has four parameters");

                // remove empty volatilities
                int[] idx = volatilities.Select((v, Index) => new { V = v, idx = Index })
                    .Where(x => x.V == 0)
                    .Select(x => x.idx)
                    .ToArray();
                
                QlArray xx = new QlArray((uint)strikes.Length - (uint)idx.Length); 
                QlArray yy = new QlArray((uint)volatilities.Length - (uint)idx.Length);
                for (uint i = 0, j = 0; i < strikes.Length; i++)
                {
                    if (volatilities[i] == 0)       // empty
                        continue;
                    xx.set(j, strikes[i]);
                    yy.set(j, volatilities[i]);
                    j++;
                }

                EndCriteria endcriteria = new EndCriteria(100000, 100, 1e-8, 1e-8, 1e-8);
                OptimizationMethod opmodel = new Simplex(0.01);
                // alpha: ATM Vol, beta: CEV param, rho: underlying/vol correlation, nu: vol of vol
                SABRInterpolation sabr = new SABRInterpolation(xx, yy, tenor, forward, 
                    initials[0], initials[1], initials[2], initials[3],
                    endcriteria, opmodel,
                    (bool)isfixed[0], (bool)isfixed[1], (bool)isfixed[2], (bool)isfixed[3], true);
                
                double err = 0;
                // only update if there exist free parameters
                if (!((bool)isfixed[0] && (bool)isfixed[1] && (bool)isfixed[2] && (bool)isfixed[3]))
                    err = sabr.update();

                // Store the option and return its id
                string id = "MODEL@" + ObjectId;
                OHRepository.Instance.storeObject(id, sabr, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }


        [ExcelFunction(Description = "Get SABR interpolated value", Category = "EliteQuantExcel - Models")]
        public static object eqModelGetSABRCalibratedParameters(
            [ExcelArgument(Description = "id of SABR model ")] string ObjectId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                SABRInterpolation option = OHRepository.Instance.getObject<SABRInterpolation>(ObjectId);

                object[,] ret = new object[6,2];
                ret[0, 0] = "alpha:";  ret[0, 1] = option.alpha();
                ret[1, 0] = "beta:"; ret[1, 1] = option.beta();
                ret[2, 0] = "nu:"; ret[2, 1] = option.nu();
                ret[3, 0] = "rho:"; ret[3, 1] = option.rho();
                ret[4, 0] = "rmsError:"; ret[4, 1] = option.rmsError();
                ret[5, 0] = "maxError:"; ret[5, 1] = option.maxError();

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Get SABR interpolated value", Category = "EliteQuantExcel - Models")]
        public static object eqModelGetSABRInterpolatedValue(
            [ExcelArgument(Description = "id of SABR model ")] string ObjectId,
            [ExcelArgument(Description = "x value ")]double x)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                SABRInterpolation option = OHRepository.Instance.getObject<SABRInterpolation>(ObjectId);
                return option.call(x);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }
        #endregion

        #region SVI
        [ExcelFunction(Description = "SVI section calibration", Category = "EliteQuantExcel - Models")]
        public static string eqModelSVI(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "ATM forward ")]double forward,
            [ExcelArgument(Description = "expiry (years of fraction ")]double tenor,
            [ExcelArgument(Description = "Strikes ")]double[] strikes,
            [ExcelArgument(Description = "implied vols ")]double[] volatilities,
            [ExcelArgument(Description = "initial parameters ")]double[] initials,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (strikes.Length != volatilities.Length)
                    throw new Exception("SABR input lengths don't match.");

                if (initials.Length != 5)
                    throw new Exception("SVI has five parameters.");

                int[] idx = volatilities.Select((v, Index) => new { V = v, idx = Index })
                    .Where(x => x.V == 0)
                    .Select(x => x.idx)
                    .ToArray();
                
                QlArray xx = new QlArray((uint)strikes.Length - (uint)idx.Length); 
                QlArray yy = new QlArray((uint)volatilities.Length - (uint)idx.Length);
                for (uint i = 0, j = 0; i < strikes.Length; i++)
                {
                    if (volatilities[i] == 0)       // empty
                        continue;
                    xx.set(j, strikes[i]);
                    yy.set(j, volatilities[i]);
                    j++;
                }

                EndCriteria endcriteria = new EndCriteria(100000, 100, 1e-8, 1e-8, 1e-8);
                OptimizationMethod opmodel = new Simplex(0.01);
                // alpha: ATM Vol, beta: CEV param, rho: underlying/vol correlation, nu: vol of vol
                SVIInterpolation svi = new SVIInterpolation(xx, yy, tenor, forward, 
                    initials[0], initials[1], initials[2], initials[3], initials[4],
                    endcriteria, opmodel, true);
                
                double err = 0;
                err = svi.update();

                // Store the option and return its id
                string id = "MODEL@" + ObjectId;
                OHRepository.Instance.storeObject(id, svi, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }


        [ExcelFunction(Description = "Get SVI interpolated value", Category = "EliteQuantExcel - Models")]
        public static object eqModelGetSVICalibratedParameters(
            [ExcelArgument(Description = "id of SVI model ")] string ObjectId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                SVIInterpolation option = OHRepository.Instance.getObject<SVIInterpolation>(ObjectId);

                object[,] ret = new object[7,2];
                ret[0, 0] = "a:";  ret[0, 1] = option.a();
                ret[1, 0] = "b:"; ret[1, 1] = option.b();
                ret[2, 0] = "rho:"; ret[2, 1] = option.rho();
                ret[3, 0] = "m:"; ret[3, 1] = option.m();
                ret[4, 0] = "sigma:"; ret[3, 1] = option.sigma();
                ret[5, 0] = "rmsError:"; ret[4, 1] = option.rmsError();
                ret[6, 0] = "maxError:"; ret[5, 1] = option.maxError();

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Get SVI interpolated value", Category = "EliteQuantExcel - Models")]
        public static object eqModelGetSVIInterpolatedValue(
            [ExcelArgument(Description = "id of SVI model ")] string ObjectId,
            [ExcelArgument(Description = "strik value ")]double x)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                SVIInterpolation option = OHRepository.Instance.getObject<SVIInterpolation>(ObjectId);
                return option.call(x);
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }
        #endregion
        
        #region DoubleExponential
        [ExcelFunction(Description = "double exponential atm vol calibration", Category = "EliteQuantExcel - Models")]
        public static string eqModelDoubleExponentialATM(
            [ExcelArgument(Description = "id of option to be constructed ")] string ObjectId,
            [ExcelArgument(Description = "expiry (years of fraction ")]double[] tenors,
            [ExcelArgument(Description = "implied vols ")]double[] volatilities,
            [ExcelArgument(Description = "initial parameters ")]double[] initials,
            [ExcelArgument(Description = "calibrate or not ")]bool calibrate,
            [ExcelArgument(Description = "parameter fixed? ")]bool issigmafixed,
            [ExcelArgument(Description = "parameter fixed? ")]bool isb1fixed,
            [ExcelArgument(Description = "parameter fixed? ")]bool isb2fixed,
            [ExcelArgument(Description = "parameter fixed? ")]bool islambdafixed,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (tenors.Length != volatilities.Length)
                    throw new Exception("double exponential input lengths don't match.");

                if (ExcelUtil.isNull(calibrate))
                    calibrate = true;
                if (ExcelUtil.isNull(issigmafixed))
                    issigmafixed = false;
                if (ExcelUtil.isNull(isb1fixed))
                    isb1fixed = false;
                if (ExcelUtil.isNull(isb2fixed))
                    isb2fixed = false;
                if (ExcelUtil.isNull(islambdafixed))
                    islambdafixed = false;
                    
                DoubleVector tt = new DoubleVector(tenors.Length);
                DoubleVector vols = new DoubleVector(volatilities.Length);

                for (uint i = 0; i < tenors.Length; i++)
                {
                    if ((ExcelUtil.isNull(tenors[i])) || (tenors[i] == 0))
                        continue;

                    tt.Add(tenors[i]);
                    vols.Add(volatilities[i]);
                }

                DoubleExponentialCalibration atmvol;
                if ((initials.Length ==4 ) && (initials != null))
                {
                    atmvol = new DoubleExponentialCalibration(tt, vols, initials[0], initials[1], initials[2], initials[3],
                        issigmafixed, isb1fixed, isb2fixed, islambdafixed); 
                }
                else
                {
                    atmvol = new DoubleExponentialCalibration(tt, vols);
                }

                if (calibrate)
                {
                    atmvol.compute();
                }

                // Store the option and return its id
                string id = "MODEL@" + ObjectId;
                OHRepository.Instance.storeObject(id, atmvol, callerAddress);
                id += "#" + (String)DateTime.Now.ToString(@"HH:mm:ss");
                return id;
                
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }


        [ExcelFunction(Description = "Get double exponential interpolated value", Category = "EliteQuantExcel - Models")]
        public static object eqModelGetDoubleExponentialCalibratedParameters(
            [ExcelArgument(Description = "id of double exponential ATM model ")] string ObjectId,
            [ExcelArgument(Description = "trigger ")]object trigger)
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                Xl.Range rng = ExcelUtil.getActiveCellRange();
                DoubleExponentialCalibration atmvol = OHRepository.Instance.getObject<DoubleExponentialCalibration>(ObjectId);

                object[,] ret = new object[6, 2];
                ret[0, 0] = "sigma:"; ret[0, 1] = atmvol.sigma();
                ret[1, 0] = "b1:"; ret[1, 1] = atmvol.b1();
                ret[2, 0] = "b2:"; ret[2, 1] = atmvol.b2();
                ret[3, 0] = "lambda:"; ret[3, 1] = atmvol.lambda();
                ret[4, 0] = "error:"; ret[4, 1] = atmvol.error();
                ret[5, 0] = "maxerror:"; ret[5, 1] = atmvol.maxError();

                return ret;
            }
            catch (Exception e)
            {
                ExcelUtil.logError(callerAddress, System.Reflection.MethodInfo.GetCurrentMethod().Name.ToString(), e.Message);
                return e.Message;
            }
        }

        [ExcelFunction(Description = "Get double exponential atm term vol", Category = "EliteQuantExcel - Models")]
        public static object eqModelGetDoubleExponentialTermVol(
            [ExcelArgument(Description = "id of double exponential atm model ")] string ObjectId,
            [ExcelArgument(Description = "from t ")] double tfrom,
            [ExcelArgument(Description = "to t ")]double tto,      // 0 to t
            [ExcelArgument(Description = "T value ")]double maturityT,
            [ExcelArgument(Description = "T Vol ")]double TQuoteVol)      // for T
        {
            if (ExcelUtil.CallFromWizard())
                return "";

            string callerAddress = "";
            callerAddress = ExcelUtil.getActiveCellAddress();

            try
            {
                if (ExcelUtil.isNull(tfrom))
                    tfrom = 0.0;

                bool bigT = true;
                if (ExcelUtil.isNull(TQuoteVol) || (TQuoteVol==0.0))
                    bigT = false;

                Xl.Range rng = ExcelUtil.getActiveCellRange();
                DoubleExponentialCalibration atmvol = OHRepository.Instance.getObject<DoubleExponentialCalibration>(ObjectId);

                double scaler = 1.0;
                if (bigT)
                {
                    DoubleVector tv = new DoubleVector(); tv.Add(maturityT);
                    DoubleVector vv = new DoubleVector(); vv.Add(TQuoteVol);

                    scaler = atmvol.k(tv, vv)[0];
                }
                    
                return scaler*atmvol.value(tfrom, tto, maturityT);
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
