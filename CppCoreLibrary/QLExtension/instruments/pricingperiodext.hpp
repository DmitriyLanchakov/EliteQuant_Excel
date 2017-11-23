#ifndef qlex_pricing_period_ext_hpp
#define qlex_pricing_period_ext_hpp

#include <instruments/dateintervalext.hpp>
#include <vector>

namespace QLExtension {

    //! Time pricingperiod described by a number of a given time unit
    /*! \ingroup datetime */
    class PricingPeriodExt : public DateIntervalExt {
      public:
		  PricingPeriodExt(const Date& startDate, const Date& endDate,
                      const Date& paymentDate, const Real quantity, 
					  const Real payCoeff = 1, const Real recCoeff = 1, 
					  const Real paySpread = 0, const Real recSpread = 0)
        : DateIntervalExt(startDate, endDate), paymentDate_(paymentDate), quantity_(quantity), 
			payCoeff_(payCoeff), recCoeff_(recCoeff), paySpread_(paySpread), recSpread_(recSpread) {
        }
        const Date paymentDate() const { return paymentDate_; }
        const Real quantity() const { return quantity_; }		// quantity for the whole period (month)

		const Real payCoeff() const { return payCoeff_; }		// payCoeff * payleg + paySpread
		const Real paySpread() const { return paySpread_; }		// payCoeff * payleg + paySpread
		const Real recCoeff() const { return recCoeff_; }		// recCoeff * recleg + recSpread
		const Real recSpread() const { return recSpread_; }		// recCoeff * recleg + recSpread

		Real getuPayDelta() const { return uPayDelta_; }
		void setuPayDelta(Real v) { uPayDelta_ = v; }
		Real getuRecDelta() const { return uRecDelta_; }
		void setuRecDelta(Real v) { uRecDelta_ = v; }

		Real getdPayDelta() const { return dPayDelta_; }
		void setdPayDelta(Real v) { dPayDelta_ = v; }
		Real getdRecDelta() const { return dRecDelta_; }
		void setdRecDelta(Real v) { dRecDelta_ = v; }

		bool getFinalized() const { return finalized_; }
		void setFinalized(bool v) { finalized_ = v; }

      private:
        Date paymentDate_;
        Real quantity_;
		Real payCoeff_;
		Real recCoeff_;
		Real paySpread_;
		Real recSpread_;

		Real uPayDelta_;
		Real dPayDelta_;
		Real uRecDelta_;
		Real dRecDelta_;
		bool finalized_;
    };

	typedef std::vector<boost::shared_ptr<PricingPeriodExt> > PricingPeriodExts;
}

#endif
