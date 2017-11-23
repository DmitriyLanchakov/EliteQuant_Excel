#ifndef qlex_svi_interpolation_hpp
#define qlex_svi_interpolation_hpp

#include <ql/utilities/null.hpp>
#include <ql/utilities/dataformatters.hpp>
#include <ql/math/interpolation.hpp>
#include <ql/math/optimization/method.hpp>
#include <ql/math/optimization/simplex.hpp>
#include <ql/math/optimization/levenbergmarquardt.hpp>
#include <ql/pricingengines/blackformula.hpp>
#include <ql/math/optimization/projectedcostfunction.hpp>
#include <ql/math/optimization/constraint.hpp>
#include <ql/math/randomnumbers/haltonrsg.hpp>

#include <termstructures/volatility/svi.hpp>

#include <boost/make_shared.hpp>
#include <boost/assign/list_of.hpp>

using namespace QuantLib;

namespace QLExtension {

namespace detail {
	class SVICoeffHolder {
	public:
		SVICoeffHolder(const Time t, const Real &forward, std::vector<Real> params)
			: t_(t), forward_(forward), params_(params),
			weights_(std::vector<Real>()), error_(Null<Real>()),
			maxError_(Null<Real>()), SVIEndCriteria_(EndCriteria::None) {
			QL_REQUIRE(t > 0.0, "expiry time must be positive: " << t
				<< " not allowed");
		}
		virtual ~SVICoeffHolder() {}

		/*! Expiry, Forward */
		Real t_;
		const Real forward_;
		/*! Parameters */
		std::vector<Real> params_;
		std::vector<Real> weights_;
		/*! Interpolation results */
		Real error_, maxError_;
		EndCriteria::Type SVIEndCriteria_;
	};

	template <class I1, class I2>
	class SVIInterpolationImpl : public Interpolation::templateImpl<I1, I2>,
								 public SVICoeffHolder
	{
	public:
		SVIInterpolationImpl(
			const I1 &xBegin, const I1 &xEnd, const I2 &yBegin, Time t,
			const Real &forward, std::vector<Real> params, bool vegaWeighted,
			const boost::shared_ptr<EndCriteria> &endCriteria,
			const boost::shared_ptr<OptimizationMethod> &optMethod,
			const Real errorAccept, const bool useMaxError, const Size maxGuesses)
			: Interpolation::templateImpl<I1, I2>(xBegin, xEnd, yBegin),
			SVICoeffHolder(t, forward, params),
			endCriteria_(endCriteria), optMethod_(optMethod),
			errorAccept_(errorAccept), useMaxError_(useMaxError),
			maxGuesses_(maxGuesses),  vegaWeighted_(vegaWeighted) {
				// if no optimization method or endCriteria is provided, we provide one
				if (!optMethod_)
					optMethod_ = boost::shared_ptr<OptimizationMethod>(
					new LevenbergMarquardt(1e-8, 1e-8, 1e-8));
				// optMethod_ = boost::shared_ptr<OptimizationMethod>(new
				//    Simplex(0.01));
				if (!endCriteria_) {
					endCriteria_ = boost::shared_ptr<EndCriteria>(
						new EndCriteria(60000, 100, 1e-8, 1e-8, 1e-8));
				}
				// equal weighted
				this->weights_ = std::vector<Real>(xEnd - xBegin, 1.0 / (xEnd - xBegin));
				completed_ = false;
			}

		// Optimization
		void update() {
			// we must update weights if it is vegaWeighted
			if (vegaWeighted_) {
				// std::vector<Real>::const_iterator x = this->xBegin_;
				// std::vector<Real>::const_iterator y = this->yBegin_;
				// std::vector<Real>::iterator w = weights_.begin();
				this->weights_.clear();
				Real weightsSum = 0.0;
				for (Size i = 0; i<Size(this->xEnd_ - this->xBegin_); ++i) {
					Real stdDev = std::sqrt((this->yBegin_[i]) * (this->yBegin_[i]) * this->t_);
					this->weights_.push_back(
						blackFormulaStdDevDerivative(this->xBegin_[i], forward_, stdDev));
					weightsSum += this->weights_.back();
				}
				// weight normalization
				std::vector<Real>::iterator w = this->weights_.begin();
				for (; w != this->weights_.end(); ++w)
					*w /= weightsSum;
			}

			try {
				SVIError costFunction(this);
				Constraint constraint = SVIConstraint();

				Array guess(params_.size());
				for (Size i = 0; i < params_.size(); i++)
				{
					guess[i] = this->params_[i];
				}
				Size iterations = 0;
				Real tmpInterpolationError;
				EndCriteria::Type tmpEndCriteria;
				do {
					Problem problem(costFunction, constraint, guess);
					tmpEndCriteria = optMethod_->minimize(problem, *endCriteria_);
					guess = problem.currentValue();
					tmpInterpolationError = useMaxError_ ? interpolationMaxError()
						: interpolationError();

					switch (tmpEndCriteria) {
					case EndCriteria::None:
					case EndCriteria::MaxIterations:
					case EndCriteria::Unknown:
						completed_ = false;
						break;
					default:
						completed_ = true;
					}
				} while (++iterations < maxGuesses_ &&
					tmpInterpolationError > errorAccept_);

				for (Size i = 0; i < guess.size(); ++i)
					this->params_[i] = guess[i];

				this->error_ = interpolationError();
				this->maxError_ = interpolationMaxError();
			}
			catch (const std::exception& e)
			{
				QL_FAIL("error in SVI calibration: " << e.what());
			}
		}

		Real value(Real x) const {
			QL_REQUIRE(x > 0.0, "strike must be positive: " << io::rate(x)
				<< " not allowed");
			return sviVolatility(x, forward_, t_, params_[0], params_[1], params_[2], params_[3], params_[4]);
		}

		Real primitive(Real) const { QL_FAIL("SVI primitive not implemented"); }
		Real derivative(Real) const { QL_FAIL("SVI derivative not implemented"); }
		Real secondDerivative(Real) const {
			QL_FAIL("SVI secondDerivative not implemented");
		}

		// calculate total squared weighted difference (L2 norm)
		Real interpolationSquaredError() const {
			Real error, totalError = 0.0;
			//std::vector<Real>::const_iterator x = this->xBegin_;
			//std::vector<Real>::const_iterator y = this->yBegin_;
			// std::vector<Real>::const_iterator w = this->weights_.begin();
			for (Size i = 0; i<Size(this->xEnd_ - this->xBegin_); ++i) {
				error = (value(this->xBegin_[i]) - this->yBegin_[i]);
				totalError += error * error * (this->weights_[i]);
			}
			return totalError;
		}

		// calculate weighted differences
		Disposable<Array> interpolationErrors(const Array &) const {
			Array results(Size(this->xEnd_ - this->xBegin_));
			//std::vector<Real>::const_iterator x = this->xBegin_; 
			//std::vector<Real>::const_iterator y = this->yBegin_;
			//std::vector<Real>::const_iterator w = this->weights_.begin();
			Array::iterator r = results.begin();
			for (Size i = 0; i < Size(this->xEnd_ - this->xBegin_); i++)
			{
				*r = (value(this->xBegin_[i]) - this->yBegin_[i]) * std::sqrt(this->weights_[i]);
			}
			return results;
		}

		Real interpolationError() const {
			Size n = this->xEnd_ - this->xBegin_;
			Real squaredError = interpolationSquaredError();
			return std::sqrt(n * squaredError / (n - 1));
		}

		Real interpolationMaxError() const {
			Real error, maxError = QL_MIN_REAL;
			I1 i = this->xBegin_;
			I2 j = this->yBegin_;
			for (; i != this->xEnd_; ++i, ++j) {
				error = std::fabs(value(*i) - *j);
				maxError = std::max(maxError, error);
			}
			return maxError;
		}
	public:
		boost::shared_ptr<EndCriteria> endCriteria_;
		boost::shared_ptr<OptimizationMethod> optMethod_;
		const Real errorAccept_;
		const bool useMaxError_;
		const Size maxGuesses_;
		bool vegaWeighted_;
		bool completed_;
		NoConstraint constraint_;

	private:
		class SVIError : public CostFunction {
		public:
			SVIError(SVIInterpolationImpl *svi) : svi_(svi) {}

			// calculate total squared weighted difference
			Real value(const Array &x) const {
				for (Size i = 0; i < svi_->params_.size(); i++)
				{
					svi_->params_[i] = x[i];
				}

				return svi_->interpolationSquaredError();
			}

			Disposable<Array> values(const Array &x) const {
				for (Size i = 0; i < svi_->params_.size(); i++)
				{
					svi_->params_[i] = x[i];
				}

				return svi_->interpolationErrors(x);
			}

		private:
			SVIInterpolationImpl *svi_;
		};

		class SVIConstraint : public Constraint {
		private:
			class Impl : public Constraint::Impl {
			public:
				bool test(const Array& params) const {
					if (params[1] < 0)
						return false;
					if (std::fabs(params[2]) >= 1)
						return false;
					if (params[4] <= 0)
						return false;
					//if (params[0] + params[1] * params[4] * std::sqrt(1-params[2]*params[2]) < 0)
					//	return false;

					return true;
				}
			};
		public:
			SVIConstraint()
				: Constraint(boost::shared_ptr<Constraint::Impl>(
				new SVIConstraint::Impl)) {}
		};
	};
}	// end of namespace detail

	//! %SVI smile interpolation between discrete volatility points.
	class SVIInterpolation : public Interpolation {
	public:
		template <class I1, class I2>
		SVIInterpolation(const I1 &xBegin,  // x = strikes
			const I1 &xEnd,
			const I2 &yBegin,  // y = volatilities
			Time t,            // option expiry
			const Real& forward,
			Real a,
			Real b,
			Real rho,
			Real m,
			Real sigma,
			bool vegaWeighted = true,
			const boost::shared_ptr<EndCriteria>& endCriteria
			= boost::shared_ptr<EndCriteria>(),
			const boost::shared_ptr<OptimizationMethod>& optMethod
			= boost::shared_ptr<OptimizationMethod>(),
			const Real errorAccept = 0.0020,
			const bool useMaxError = false,
			const Size maxGuesses = 1) {

			impl_ = boost::shared_ptr<Interpolation::Impl>(
				new detail::SVIInterpolationImpl<I1, I2>(
				xBegin, xEnd, yBegin, t, forward,
				boost::assign::list_of(a)(b)(rho)(m)(sigma),
				vegaWeighted, endCriteria, optMethod, errorAccept, useMaxError,
				maxGuesses));
			coeffs_ = boost::dynamic_pointer_cast<
				detail::SVICoeffHolder >(impl_);
		}
		Real expiry()  const { return coeffs_->t_; }
		Real forward() const { return coeffs_->forward_; }
		Real a()   const { return coeffs_->params_[0]; }
		Real b()    const { return coeffs_->params_[1]; }
		Real rho()      const { return coeffs_->params_[2]; }
		Real m()     const { return coeffs_->params_[3]; }
		Real sigma()     const { return coeffs_->params_[4]; }
		Real rmsError() const { return coeffs_->error_; }
		Real maxError() const { return coeffs_->maxError_; }
		const std::vector<Real>& interpolationWeights() const {
			return coeffs_->weights_;
		}
		EndCriteria::Type endCriteria() { return coeffs_->SVIEndCriteria_; }

	private:
		boost::shared_ptr<detail::SVICoeffHolder> coeffs_;
	};
}	// end of namespace QLExtension

#endif
