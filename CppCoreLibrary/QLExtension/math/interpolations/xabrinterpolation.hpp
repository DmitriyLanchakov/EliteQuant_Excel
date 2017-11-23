#ifndef qlex_xabr_interpolation_hpp
#define qlex_xabr_interpolation_hpp

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

using namespace QuantLib;

namespace QLExtension {

namespace detail {

template <typename Model> class XABRCoeffHolder {
  public:
    XABRCoeffHolder(const Time t, const Real &forward, std::vector<Real> params,
                    std::vector<bool> paramIsFixed)
        : t_(t), forward_(forward), params_(params),
          paramIsFixed_(paramIsFixed.size(), false),
          weights_(std::vector<Real>()), error_(Null<Real>()),
          maxError_(Null<Real>()), XABREndCriteria_(EndCriteria::None) {
        QL_REQUIRE(t > 0.0, "expiry time must be positive: " << t
                                                             << " not allowed");
        QL_REQUIRE(params.size() == Model().dimension(),
                   "wrong number of parameters (" << params.size()
                                                  << "), should be "
                                                  << Model().dimension());
        QL_REQUIRE(paramIsFixed.size() == Model().dimension(),
                   "wrong number of fixed parameters flags ("
                       << paramIsFixed.size() << "), should be "
                       << Model().dimension());

        for (Size i = 0; i < params.size(); ++i) {
            if (params[i] != Null<Real>())
                paramIsFixed_[i] = paramIsFixed[i];
        }
        Model().defaultValues(params_, paramIsFixed_, forward_, t_);
        updateModelInstance();
    }
    virtual ~XABRCoeffHolder() {}

    void updateModelInstance() {
        // forward might have changed
        QL_REQUIRE(forward_ > 0.0,
                   "forward must be positive: " << forward_ << " not allowed");
        modelInstance_ = Model().instance(t_, forward_, params_);
    }

    /*! Expiry, Forward */
    Real t_;
    const Real forward_;
    /*! Parameters */
    std::vector<Real> params_;
    std::vector<bool> paramIsFixed_;
    std::vector<Real> weights_;
    /*! Interpolation results */
    Real error_, maxError_;
    EndCriteria::Type XABREndCriteria_;
    /*! Model instance (if required) */
    boost::shared_ptr<typename Model::type> modelInstance_;
};

template <class I1, class I2, typename Model>
class XABRInterpolationImpl : public Interpolation::templateImpl<I1, I2>,
                              public XABRCoeffHolder<Model> {
  public:
    XABRInterpolationImpl(
        const I1 &xBegin, const I1 &xEnd, const I2 &yBegin, Time t,
        const Real &forward, std::vector<Real> params,
        std::vector<bool> paramIsFixed, bool vegaWeighted,
        const boost::shared_ptr<EndCriteria> &endCriteria,
        const boost::shared_ptr<OptimizationMethod> &optMethod,
        const Real errorAccept, const bool useMaxError, const Size maxGuesses)
        : Interpolation::templateImpl<I1, I2>(xBegin, xEnd, yBegin),
          XABRCoeffHolder<Model>(t, forward, params, paramIsFixed),
          endCriteria_(endCriteria), optMethod_(optMethod),
          errorAccept_(errorAccept), useMaxError_(useMaxError),
          maxGuesses_(maxGuesses), forward_(forward),
          vegaWeighted_(vegaWeighted) {
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
        this->weights_ = std::vector<Real>(xEnd - xBegin, 1.0 / (xEnd - xBegin));
    }

    void update() {

        this->updateModelInstance();

        // we should also check that y contains positive values only

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

        // there is nothing to optimize
        if (std::accumulate(this->paramIsFixed_.begin(),
                            this->paramIsFixed_.end(), true,
                            std::logical_and<bool>())) {
            this->error_ = interpolationError();
            this->maxError_ = interpolationMaxError();
            this->XABREndCriteria_ = EndCriteria::None;
            return;
        } else {
            XABRError costFunction(this);

            Array guess(Model().dimension());
            for (Size i = 0; i < guess.size(); ++i)
                guess[i] = this->params_[i];

            Size iterations = 0;
            Size freeParameters = 0;
            Real bestError = QL_MAX_REAL;
            Array bestParameters;
            for (Size i = 0; i < Model().dimension(); ++i)
                if (!this->paramIsFixed_[i])
                    ++freeParameters;
            HaltonRsg halton(freeParameters, 42);
            EndCriteria::Type tmpEndCriteria;
            Real tmpInterpolationError;

            do {

                if (iterations > 0) {
                    HaltonRsg::sample_type s = halton.nextSequence();
                    Model().guess(guess, this->paramIsFixed_, forward_, this->t_, s.value);
                    for (Size i = 0; i < this->paramIsFixed_.size(); ++i)
                        if (this->paramIsFixed_[i])
                            guess[i] = this->params_[i];
                }

                Array inversedTransformatedGuess(Model().inverse(
                    guess, this->paramIsFixed_, this->params_, forward_));

                ProjectedCostFunction constrainedXABRError(
                    costFunction, inversedTransformatedGuess, this->paramIsFixed_);

                Array projectedGuess(
                    constrainedXABRError.project(inversedTransformatedGuess));

                NoConstraint constraint;
                Problem problem(constrainedXABRError, constraint,
                                projectedGuess);
                tmpEndCriteria = optMethod_->minimize(problem, *endCriteria_);
                Array projectedResult(problem.currentValue());
                Array transfResult(
                    constrainedXABRError.include(projectedResult));

                Array result = Model().direct(transfResult, this->paramIsFixed_,
                                              this->params_, forward_);
                tmpInterpolationError = useMaxError_ ? interpolationMaxError()
                                                     : interpolationError();

                if (tmpInterpolationError < bestError) {
                    bestError = tmpInterpolationError;
                    bestParameters = result;
                    this->XABREndCriteria_ = tmpEndCriteria;
                }

            } while (++iterations < maxGuesses_ &&
                     tmpInterpolationError > errorAccept_);

            for (Size i = 0; i < bestParameters.size(); ++i)
                this->params_[i] = bestParameters[i];

            this->error_ = interpolationError();
            this->maxError_ = interpolationMaxError();
        }
    }

    Real value(Real x) const {
        QL_REQUIRE(x > 0.0, "strike must be positive: " << io::rate(x)
                                                        << " not allowed");
        return this->modelInstance_->volatility(x);
    }

    Real primitive(Real) const { QL_FAIL("XABR primitive not implemented"); }
    Real derivative(Real) const { QL_FAIL("XABR derivative not implemented"); }
    Real secondDerivative(Real) const {
        QL_FAIL("XABR secondDerivative not implemented");
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

  private:
    class XABRError : public CostFunction {
      public:
        XABRError(XABRInterpolationImpl *xabr) : xabr_(xabr) {}

        Real value(const Array &x) const {
            const Array y = Model().direct(x, xabr_->paramIsFixed_,
                                           xabr_->params_, xabr_->forward_);
            for (Size i = 0; i <xabr_-> params_.size(); ++i)
                xabr_->params_[i] = y[i];
            xabr_->updateModelInstance();
            return xabr_->interpolationSquaredError();
        }

        Disposable<Array> values(const Array &x) const {
            const Array y = Model().direct(x, xabr_->paramIsFixed_,
                                           xabr_->params_, xabr_->forward_);
            for (Size i = 0; i < xabr_->params_.size(); ++i)
                xabr_->params_[i] = y[i];
            xabr_->updateModelInstance();
            return xabr_->interpolationErrors(x);
        }

      private:
        XABRInterpolationImpl *xabr_;
    };
    boost::shared_ptr<EndCriteria> endCriteria_;
    boost::shared_ptr<OptimizationMethod> optMethod_;
    const Real errorAccept_;
    const bool useMaxError_;
    const Size maxGuesses_;
    const Real forward_;
    bool vegaWeighted_;
    NoConstraint constraint_;
};

} // namespace detail
} // namespace QLExtension

#endif
