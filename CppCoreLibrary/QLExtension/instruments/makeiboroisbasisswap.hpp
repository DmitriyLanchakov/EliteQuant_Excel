#ifndef qlex_makeiboroisbasisswap_hpp
#define qlex_makeiboroisbasisswap_hpp

#include <instruments/iboroisbasisswap.hpp>
#include <ql/time/dategenerationrule.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>

using namespace QuantLib;

namespace QLExtension {

	//! helper class
	/*! This class provides a more comfortable way
	to instantiate ibor vs. overnight indexed swaps.
	*/
	class MakeIBOROISBasisSwap {
	public:
		MakeIBOROISBasisSwap(const Period& swapTenor,
			const boost::shared_ptr<IborIndex>& iborIndex,
			const boost::shared_ptr<OvernightIndex>& overnightIndex,
			Rate spread = Null<Rate>(),
			const Period& fwdStart = 0 * Days);

		operator IBOROISBasisSwap() const;
		operator boost::shared_ptr<IBOROISBasisSwap>() const;

		MakeIBOROISBasisSwap& withType(IBOROISBasisSwap::Type type);
		MakeIBOROISBasisSwap& withNominal(Real n);
		MakeIBOROISBasisSwap& withSettlementDays(Natural fixingDays);
		MakeIBOROISBasisSwap& withEffectiveDate(const Date&);
		MakeIBOROISBasisSwap& withTerminationDate(const Date&);
		MakeIBOROISBasisSwap& withEndOfMonth(bool flag = true);
		MakeIBOROISBasisSwap& withPaymentConvention(BusinessDayConvention bc);

		MakeIBOROISBasisSwap& withFloatingLegTenor(const Period& t);
		MakeIBOROISBasisSwap& withFloatingLegCalendar(const Calendar& cal);
		MakeIBOROISBasisSwap& withFloatingLegConvention(BusinessDayConvention bdc);
		MakeIBOROISBasisSwap& withFloatingLegTerminationDateConvention(
			BusinessDayConvention bdc);
		MakeIBOROISBasisSwap& withFloatingLegRule(DateGeneration::Rule r);
		MakeIBOROISBasisSwap& withFloatingLegDayCount(const DayCounter& dc);

		MakeIBOROISBasisSwap& withOvernightLegTenor(const Period& t);
		MakeIBOROISBasisSwap& withOvernightLegCalendar(const Calendar& cal);
		MakeIBOROISBasisSwap& withOvernightLegConvention(BusinessDayConvention bdc);
		MakeIBOROISBasisSwap& withOvernightLegTerminationDateConvention(
			BusinessDayConvention bdc);
		MakeIBOROISBasisSwap& withOvernightLegRule(DateGeneration::Rule r);
		MakeIBOROISBasisSwap& withOvernightLegDayCount(const DayCounter& dc);
		MakeIBOROISBasisSwap& withOvernightLegSpread(Spread sp);

		MakeIBOROISBasisSwap& withDiscountingTermStructure(
			const Handle<YieldTermStructure>& discountingTermStructure);

	private:
		Period swapTenor_;
		boost::shared_ptr<IborIndex> floatingIndex_;
		boost::shared_ptr<OvernightIndex> overnightIndex_;
		Spread overnightSpread_;
		Period forwardStart_;
		bool endOfMonth_;

		IBOROISBasisSwap::Type type_;
		Real nominal_;
		Natural fixingDays_;
		Date effectiveDate_, terminationDate_;
		BusinessDayConvention paymentConvention_;

		Period floatingLegTenor_;
		Calendar floatingLegCalendar_;
		BusinessDayConvention floatingLegConvention_;
		BusinessDayConvention floatingLegTerminationDateConvention_;
		DateGeneration::Rule floatingLegRule_;
		DayCounter floatingLegDayCount_;

		Period overnightLegTenor_;
		Calendar overnightLegCalendar_;
		BusinessDayConvention overnightLegConvention_;
		BusinessDayConvention overnightLegTerminationDateConvention_;
		DateGeneration::Rule overnightLegRule_;
		DayCounter overnightLegDayCount_;

		boost::shared_ptr<PricingEngine> engine_;
	};
}

#endif
