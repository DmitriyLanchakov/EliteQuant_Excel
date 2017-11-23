// simultabeous bootstrap
// https://github.com/kosynski/quantlib
#ifndef qlex_oisbasisratehelper_hpp
#define qlex_oisbasisratehelper_hpp

#include <instruments/iboroisbasisswap.hpp>
#include <instruments/genericswap.hpp>
#include <cashflows/overnightindexedcoupon.hpp>
#include <ql/termstructures/yield/ratehelpers.hpp>
#include <ql/instruments/overnightindexedswap.hpp>

using namespace QuantLib;

namespace QLExtension {

	//! Rate helper for bootstrapping over Ibor vs. Overnight Indexed basis Swap rates
	class IBOROISBasisRateHelper : public RelativeDateRateHelper {
	public:
		IBOROISBasisRateHelper(Natural settlementDays,
			const Period& tenor, // swap maturity
			const Handle<Quote>& overnightSpread,
			const boost::shared_ptr<IborIndex>& iborIndex,
			const boost::shared_ptr<OvernightIndex>& overnightIndex,
			// exogenous discounting curve
			const Handle<YieldTermStructure>& discountingCurve
			= Handle<YieldTermStructure>());
		//! \name RateHelper interface
		//@{
		Real impliedQuote() const;
		void setTermStructure(YieldTermStructure*);
		//@}
		//! \name inspectors
		//@{
		boost::shared_ptr<IBOROISBasisSwap> swap() const { return swap_; }
		//@}
		//! \name Visitability
		//@{
		void accept(AcyclicVisitor&);
		//@}
	protected:
		void initializeDates();

		Natural settlementDays_;
		Period tenor_;
		boost::shared_ptr<IborIndex> iborIndex_;
		boost::shared_ptr<OvernightIndex> overnightIndex_;

		boost::shared_ptr<IBOROISBasisSwap> swap_;
		RelinkableHandle<YieldTermStructure> termStructureHandle_;

		Handle<YieldTermStructure> discountHandle_;
		RelinkableHandle<YieldTermStructure> discountRelinkableHandle_;
	};

	//! Rate helper for bootstrapping over Fixed vs. Overnight Indexed basis Swap rates
	class FixedOISBasisRateHelper : public RelativeDateRateHelper {
	public:
		FixedOISBasisRateHelper(Natural settlementDays,
			const Period& tenor, // swap maturity
			const Handle<Quote>& overnightSpread,
			const Handle<Quote>& fixedRate,
			Frequency fixedFrequency,
			BusinessDayConvention fixedConvention,
			const DayCounter& fixedDayCount,
			const boost::shared_ptr<OvernightIndex>& overnightIndex,
			Frequency overnightFrequency,
			// exogenous discounting curve
			const Handle<YieldTermStructure>& discountingCurve
			= Handle<YieldTermStructure>());
		//! \name RateHelper interface
		//@{
		Real impliedQuote() const;
		void setTermStructure(YieldTermStructure*);
		//@}
		//! \name inspectors
		//@{
		boost::shared_ptr<Swap> swap() const { return swap_; }
		//@}
		//! \name Visitability
		//@{
		void accept(AcyclicVisitor&);
		//@}
		//! \name Observer interface
		//@{
		void update();
		//@}
	protected:
		void initializeDates();

		Natural settlementDays_;
		Period tenor_;
		Handle<Quote> fixedRate_;
		Real usedFixedRate_;
		Frequency fixedFrequency_;
		BusinessDayConvention fixedConvention_;
		DayCounter fixedDayCount_;
		boost::shared_ptr<OvernightIndex> overnightIndex_;
		Frequency overnightFrequency_;

		boost::shared_ptr<Swap> swap_;
		RelinkableHandle<YieldTermStructure> termStructureHandle_;

		Handle<YieldTermStructure> discountHandle_;
		RelinkableHandle<YieldTermStructure> discountRelinkableHandle_;
	};

	// Libor basis swap
	class IBORBasisRateHelper : public RelativeDateRateHelper {
	public:
		IBORBasisRateHelper(Natural settlementDays,
			const Period& tenor, // swap maturity
			const Handle<Quote>& basis,
			const boost::shared_ptr<IborIndex>& baseLegIborIndex,
			const boost::shared_ptr<IborIndex>& basisLegIborIndex, 
			// exogenous discounting curve
			const Handle<YieldTermStructure>& discountingCurve
			= Handle<YieldTermStructure>());
		//! \name RateHelper interface
		//@{
		Real impliedQuote() const;
		void setTermStructure(YieldTermStructure*);
		//@}
		//! \name inspectors
		//@{
		boost::shared_ptr<GenericSwap> swap() const { return swap_; }
		//@}
		//! \name Visitability
		//@{
		void accept(AcyclicVisitor&);
		//@}
		//! \name Observer interface
		//@{
		void update();
		//@}
	protected:
		void initializeDates();

		Natural settlementDays_;
		Period tenor_;
		boost::shared_ptr<IborIndex> baseLegIndex_;
		boost::shared_ptr<IborIndex> basisLegIndex_;

		boost::shared_ptr<GenericSwap> swap_;
		RelinkableHandle<YieldTermStructure> termStructureHandle_;

		Handle<YieldTermStructure> discountHandle_;
		RelinkableHandle<YieldTermStructure> discountRelinkableHandle_;
	};
}

#endif
