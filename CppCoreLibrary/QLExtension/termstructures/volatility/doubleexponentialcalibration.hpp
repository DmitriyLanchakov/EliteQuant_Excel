#ifndef qlex_doubleexponentialcalibration_hpp
#define qlex_doubleexponentialcalibration_hpp

#include <ql/math/optimization/endcriteria.hpp>
#include <ql/math/optimization/projectedcostfunction.hpp>
#include <ql/math/array.hpp>
#include <ql/quote.hpp>
#include <boost/shared_ptr.hpp>
#include <vector>

using namespace QuantLib;

namespace QLExtension {
    
    class QuantLib::Quote;
    class QuantLib::OptimizationMethod;
    class QuantLib::ParametersTransformation;

	// Time homogeneous vol term structure
	// pp. 104 Valuation and Risk Management in Energy Market
	// sigma^2(T-t) = simga^2[exp(-2b_1(T-1)+lambda*exp(-2b_2(T-t)]
    class DoubleExponentialCalibration {
    
    private:

		class DoubleExponentialError : public CostFunction {
          public:
            DoubleExponentialError(DoubleExponentialCalibration* dblexp) : dblexp_(dblexp) {}

            Real value(const Array& x) const {
				dblexp_->sigma_ = x[0];
				dblexp_->b1_ = x[1];
				dblexp_->b2_ = x[2];
				dblexp_->lambda_ = x[3];
				return dblexp_->error();
            }
            Disposable<Array> values(const Array& x) const {
				dblexp_->sigma_ = x[0];
				dblexp_->b1_ = x[1];
				dblexp_->b2_ = x[2];
				dblexp_->lambda_ = x[3];
				return dblexp_->errors();
            }
          private:
            DoubleExponentialCalibration* dblexp_;
        };

      public:
        DoubleExponentialCalibration() {};
		DoubleExponentialCalibration(
             const std::vector<Real>& t,
             const std::vector<Real>& blackVols,
             Real sigmaGuess = 0.10,
             Real b1Guess =  0.25,
             Real b2Guess =  0.60,
             Real lambdaGuess =  0.5,
             bool sigmaIsFixed = false,
             bool b1IsFixed = false,
             bool b2IsFixed = false,
             bool lambdaIsFixed = false,
             bool vegaWeighted = false,
             const boost::shared_ptr<EndCriteria>& endCriteria
                      = boost::shared_ptr<EndCriteria>(),
             const boost::shared_ptr<OptimizationMethod>& method
                      = boost::shared_ptr<OptimizationMethod>());

        //! adjustment factors needed to match Black vols
		std::vector<Real> k() const;		// big T approach
        std::vector<Real> k(const std::vector<Real>& t,
                            const std::vector<Real>& blackVols) const;
        void compute();
        //calibration results
		// obsolute
        Real value(Real x, Real T) const;
		Real value(Real t1, Real t2, Real T) const;
		
		// max abs(modelvol - marketvol)
        Real maxError() const;
		// vector (modelvol-marketvol)*sqrt(weight)
        Disposable<Array> errors() const;
		// square sum of errors()
		Real error() const;

        EndCriteria::Type endCriteria() const;

        Real sigma() const;
        Real b1() const;
        Real b2() const;
        Real lambda() const;

        bool sigmaIsFixed_, b1IsFixed_, b2IsFixed_, lambdaIsFixed_;
        Real sigma_, b1_, b2_, lambda_;

      private:

        // optimization method used for fitting
        mutable EndCriteria::Type dblexpEndCriteria_;
        boost::shared_ptr<EndCriteria> endCriteria_;
        boost::shared_ptr<OptimizationMethod> optMethod_;
        mutable std::vector<Real> weights_;
        bool vegaWeighted_;
        //! Parameters
        std::vector<Real> times_, blackVols_;

		inline void validateDoubleExponentialParameters(Real sigma,
			Real b1, // no condition on b
			Real b2,
			Real lambda) {
			QL_REQUIRE(sigma > 0,
				"sigma (" << sigma << ") must be positive");
			QL_REQUIRE(b1 > 0,
				"b1 (" << b1 << ") must be positive");
			QL_REQUIRE(b2 > 0,
				"b2 (" << b2 << ") must be positive");
			QL_REQUIRE(lambda > 0,
				"lambda (" << lambda << ") must be positive");
		}
    };
}

#endif
