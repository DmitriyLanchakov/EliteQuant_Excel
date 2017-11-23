#include <ql/math/optimization/method.hpp>
#include <ql/math/optimization/constraint.hpp>
#include <ql/math/optimization/levenbergmarquardt.hpp>
#include <ql/pricingengines/blackformula.hpp>
#include <ql/termstructures/volatility/abcd.hpp>
#include <ql/math/distributions/normaldistribution.hpp>
#include <ql/math/interpolations/abcdinterpolation.hpp>

#include <termstructures/volatility/doubleexponentialcalibration.hpp>

namespace QLExtension {

	DoubleExponentialCalibration::DoubleExponentialCalibration(
               const std::vector<Real>& t,
               const std::vector<Real>& blackVols,
               Real sigma, Real b1, Real b2, Real lambda,
               bool sigmaIsFixed, bool b1IsFixed, bool b2IsFixed, bool dIsFixed,
               bool vegaWeighted,
               const boost::shared_ptr<EndCriteria>& endCriteria,
               const boost::shared_ptr<OptimizationMethod>& optMethod)
    : sigmaIsFixed_(sigmaIsFixed), b1IsFixed_(b1IsFixed),
      b2IsFixed_(b2IsFixed), lambdaIsFixed_(dIsFixed),
      sigma_(sigma), b1_(b1), b2_(b2), lambda_(lambda),
	  dblexpEndCriteria_(EndCriteria::None), endCriteria_(endCriteria),
      optMethod_(optMethod), weights_(blackVols.size(), 1.0/blackVols.size()),
      vegaWeighted_(vegaWeighted),
      times_(t), blackVols_(blackVols) {

        QL_REQUIRE(blackVols.size()==t.size(),
                       "mismatch between number of times (" << t.size() <<
                       ") and blackVols (" << blackVols.size() << ")");

        // if no optimization method or endCriteria is provided, we provide one
        if (!optMethod_)
            optMethod_ = boost::shared_ptr<OptimizationMethod>(new
                LevenbergMarquardt(1e-8, 1e-8, 1e-8));
            //method_ = boost::shared_ptr<OptimizationMethod>(new
            //    Simplex(0.01));
        if (!endCriteria_)
            //endCriteria_ = boost::shared_ptr<EndCriteria>(new
            //    EndCriteria(60000, 100, 1e-8, 1e-8, 1e-8));
            endCriteria_ = boost::shared_ptr<EndCriteria>(new
                //EndCriteria(1000, 100, 1.0e-8, 0.3e-4, 0.3e-4));   // Why 0.3e-4 ?
				EndCriteria(50000, 5000, 1.0e-8, 1e-6, 1e-6));   // Why 0.3e-4 ?
    }

	void DoubleExponentialCalibration::compute() {
        if (vegaWeighted_) {
            Real weightsSum = 0.0;
            for (Size i=0; i<times_.size() ; i++) {
                Real stdDev = std::sqrt(blackVols_[i]* blackVols_[i]* times_[i]);
                // when strike==forward, the blackFormulaStdDevDerivative becomes
                weights_[i] = CumulativeNormalDistribution().derivative(.5*stdDev);
                weightsSum += weights_[i];
            }
            // weight normalization
            for (Size i=0; i<times_.size() ; i++) {
                weights_[i] /= weightsSum;
            }
        }

        // there is nothing to optimize
        if (sigmaIsFixed_ && b1IsFixed_ && b2IsFixed_ && lambdaIsFixed_) {
			dblexpEndCriteria_ = EndCriteria::None;
            //error_ = interpolationError();
            //maxError_ = interpolationMaxError();
            return;
        } else {

			DoubleExponentialError costFunction(this);

            Array guess(4);
            guess[0] = sigma_;
            guess[1] = b1_;
            guess[2] = b2_;
            guess[3] = lambda_;

            std::vector<bool> parameterAreFixed(4);
            parameterAreFixed[0] = sigmaIsFixed_;
            parameterAreFixed[1] = b1IsFixed_;
            parameterAreFixed[2] = b2IsFixed_;
            parameterAreFixed[3] = lambdaIsFixed_;

            ProjectedCostFunction projectedDoubleExponentialCostFunction(costFunction,
									guess, parameterAreFixed);

            Array projectedGuess
				(projectedDoubleExponentialCostFunction.project(guess));

            // NoConstraint constraint;
			//PositiveConstraint constraint;
			BoundaryConstraint constraint(0.0, 1.0);
			Problem problem(projectedDoubleExponentialCostFunction, constraint, projectedGuess);
			dblexpEndCriteria_ = optMethod_->minimize(problem, *endCriteria_);
            Array projectedResult(problem.currentValue());
			Array result(projectedDoubleExponentialCostFunction.include(projectedResult));

            sigma_ = result[0];
            b1_ = result[1];
            b2_ = result[2];
            lambda_ = result[3];

            validateDoubleExponentialParameters(sigma_, b1_, b2_, lambda_);
        }
    }

	Real DoubleExponentialCalibration::sigma() const {
        return sigma_;
    }

	Real DoubleExponentialCalibration::b1() const {
        return b1_;
    }

	Real DoubleExponentialCalibration::b2() const {
        return b2_;
    }

	Real DoubleExponentialCalibration::lambda() const {
        return lambda_;
    }

	// term vol between 0 and x for T forwards
	// obsolete
	Real DoubleExponentialCalibration::value(Real x, Real T) const {
		Real a1 = 0.5 / b1_*(std::exp(-2 * b1_*(T - x)) - std::exp(-2 * b1_*T));
		Real a2 = 0.5 * lambda_ / b2_*(std::exp(-2 * b2_*(T - x)) - std::exp(-2 * b2_*T));
		return std::sqrt((a1+a2)/x) * sigma_;
    }

	// term vol between t1 and t2 for T forwards
	Real DoubleExponentialCalibration::value(Real t1, Real t2, Real T) const {
		Real a1 = (b1_ == 0) ? 0.5*(t2-t1) : 0.5 / b1_*(std::exp(-2 * b1_*(T - t2)) - std::exp(-2 * b1_*(T - t1)));
		Real a2 = (b2_ == 0) ? 0.5*(t2-t1) : 0.5 * lambda_ / b2_*(std::exp(-2 * b2_*(T - t2)) - std::exp(-2 * b2_*(T - t1)));
		return std::sqrt((a1 + a2) / (t2-t1)) * sigma_;
	}

	std::vector<Real> DoubleExponentialCalibration::k() const {
		return k(times_, blackVols_);
	}

	std::vector<Real> DoubleExponentialCalibration::k(const std::vector<Real>& t,
                                         const std::vector<Real>& blackVols) const {
        QL_REQUIRE(blackVols.size()==t.size(),
               "mismatch between number of times (" << t.size() <<
               ") and blackVols (" << blackVols.size() << ")");
        std::vector<Real> k(t.size());
        for (Size i=0; i<t.size() ; i++) {
            k[i]=blackVols[i]/value(0, t[i],t[i]);
        }
        return k;
    }

	Real DoubleExponentialCalibration::error() const {
        Size n = times_.size();
        Real error, squaredError = 0.0;
        for (Size i=0; i<times_.size() ; i++) {
            error = (value(0, times_[i],times_[i]) - blackVols_[i]);
            squaredError += error * error *(weights_[i]);
        }
        return std::sqrt(n*squaredError/(n-1));
    }

	Real DoubleExponentialCalibration::maxError() const {
        Real error, maxError = QL_MIN_REAL;
        for (Size i=0; i<times_.size() ; i++) {
            error = std::fabs(value(0, times_[i], times_[i]) - blackVols_[i]);
            maxError = std::max(maxError, error);
        }
        return maxError;
    }

    // calculate weighted differences
	Disposable<Array> DoubleExponentialCalibration::errors() const {
        Array results(times_.size());
        for (Size i=0; i<times_.size() ; i++) {
            results[i] = (value(0, times_[i], times_[i]) - blackVols_[i])* std::sqrt(weights_[i]);
        }
        return results;
    }

	EndCriteria::Type DoubleExponentialCalibration::endCriteria() const{
		return dblexpEndCriteria_;
    }

}
