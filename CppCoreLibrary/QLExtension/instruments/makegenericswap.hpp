#ifndef qlex_makegenericswap_hpp
#define qlex_makegenericswap_hpp

#include <instruments/genericswap.hpp>
#include <ql/time/dategenerationrule.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>

using namespace QuantLib;

namespace QLExtension {

	//! helper class
	/*! This class provides a more comfortable way
	to instantiate ibor vs. libor basis indexed swaps.
	*/
	class MakeGenericSwap {
	public:
		MakeGenericSwap(const Period& swapTenor,
			const boost::shared_ptr<IborIndex>& baseLegIndex,
			const boost::shared_ptr<IborIndex>& basisLegIndex,
			Rate spread = Null<Rate>(),
			const Period& fwdStart = 0 * Days);

		operator GenericSwap() const;
		operator boost::shared_ptr<GenericSwap>() const;

		MakeGenericSwap& withType(GenericSwap::Type type);
		MakeGenericSwap& withNominal(Real n);
		MakeGenericSwap& withSettlementDays(Natural fixingDays);
		MakeGenericSwap& withEffectiveDate(const Date&);
		MakeGenericSwap& withTerminationDate(const Date&);
		MakeGenericSwap& withEndOfMonth(bool flag = true);
		MakeGenericSwap& withPaymentConvention(BusinessDayConvention bc);

		MakeGenericSwap& withBaseLegTenor(const Period& t);
		MakeGenericSwap& withBaseLegCalendar(const Calendar& cal);
		MakeGenericSwap& withBaseLegConvention(BusinessDayConvention bdc);
		MakeGenericSwap& withBaseLegTerminationDateConvention(
			BusinessDayConvention bdc);
		MakeGenericSwap& withBaseLegRule(DateGeneration::Rule r);
		MakeGenericSwap& withBaseLegDayCount(const DayCounter& dc);

		MakeGenericSwap& withBasisLegTenor(const Period& t);
		MakeGenericSwap& withBasisLegCalendar(const Calendar& cal);
		MakeGenericSwap& withBasisLegConvention(BusinessDayConvention bdc);
		MakeGenericSwap& withBasisLegTerminationDateConvention(
			BusinessDayConvention bdc);
		MakeGenericSwap& withBasisLegRule(DateGeneration::Rule r);
		MakeGenericSwap& withBasisLegDayCount(const DayCounter& dc);
		MakeGenericSwap& withBasisLegSpread(Spread sp);

		MakeGenericSwap& withDiscountingTermStructure(
			const Handle<YieldTermStructure>& discountingTermStructure);

	private:
		Period swapTenor_;
		boost::shared_ptr<IborIndex> baseLegIndex_;
		boost::shared_ptr<IborIndex> basisLegIndex_;
		Spread spread_;
		Period forwardStart_;
		bool endOfMonth_;

		GenericSwap::Type type_;
		Real nominal_;
		Natural fixingDays_;
		Date effectiveDate_, terminationDate_;
		BusinessDayConvention paymentConvention_;

		Period baseLegTenor_;
		Calendar baseLegCalendar_;
		BusinessDayConvention baseLegConvention_;
		BusinessDayConvention baseLegTerminationDateConvention_;
		DateGeneration::Rule baseLegRule_;
		DayCounter baseLegDayCount_;

		Period basisLegTenor_;
		Calendar basisLegCalendar_;
		BusinessDayConvention basisLegConvention_;
		BusinessDayConvention basisLegTerminationDateConvention_;
		DateGeneration::Rule basisLegRule_;
		DayCounter basisLegDayCount_;

		boost::shared_ptr<PricingEngine> engine_;
	};
}

#endif
