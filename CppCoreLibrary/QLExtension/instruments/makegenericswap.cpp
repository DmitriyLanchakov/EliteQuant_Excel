#include <instruments/MakeGenericSwap.hpp>
#include <ql/pricingengines/swap/discountingswapengine.hpp>
#include <ql/time/schedule.hpp>
#include <ql/indexes/iborindex.hpp>

namespace QLExtension {

	MakeGenericSwap::MakeGenericSwap(const Period& swapTenor,
		const boost::shared_ptr<IborIndex>& baseLegIndex,
		const boost::shared_ptr<IborIndex>& basisLegIndex,
		Rate spread,
		const Period& fwdStart)
		: swapTenor_(swapTenor), baseLegIndex_(baseLegIndex), basisLegIndex_(basisLegIndex),
		spread_(spread), forwardStart_(fwdStart),
		type_(GenericSwap::Payer), nominal_(1.0), fixingDays_(2),
		effectiveDate_(Date()), terminationDate_(Date()),
		paymentConvention_(ModifiedFollowing),
		endOfMonth_(1 * Months <= swapTenor && swapTenor <= 2 * Years ? true : false),
		baseLegTenor_(baseLegIndex->tenor()),
		baseLegCalendar_(baseLegIndex->fixingCalendar()),
		baseLegConvention_(baseLegIndex->businessDayConvention()),
		baseLegTerminationDateConvention_(baseLegIndex->businessDayConvention()),
		baseLegRule_(DateGeneration::Backward),
		baseLegDayCount_(baseLegIndex->dayCounter()),
		basisLegTenor_(basisLegIndex->tenor()),//exchange at the end of each ibor period
		basisLegCalendar_(basisLegIndex->fixingCalendar()),
		basisLegConvention_(basisLegIndex->businessDayConvention()),
		basisLegTerminationDateConvention_(basisLegIndex->businessDayConvention()),
		basisLegRule_(DateGeneration::Backward),
		basisLegDayCount_(basisLegIndex->dayCounter()),
		engine_(new DiscountingSwapEngine(baseLegIndex->forwardingTermStructure())){}

	MakeGenericSwap::operator GenericSwap() const {
		boost::shared_ptr<GenericSwap> oisbasis = *this;
		return *oisbasis;
	}

	MakeGenericSwap::operator boost::shared_ptr<GenericSwap>() const {

		const Calendar& calendar = basisLegIndex_->fixingCalendar();

		Date startDate;
		if (effectiveDate_ != Date())
			startDate = effectiveDate_;
		else {
			Date referenceDate = Settings::instance().evaluationDate();
			Date spotDate = calendar.advance(referenceDate,
				fixingDays_*Days);
			startDate = spotDate + forwardStart_;
		}

		Date endDate;
		if (terminationDate_ != Date()) {
			endDate = terminationDate_;
		}
		else {
			if (endOfMonth_) {
				endDate = calendar.advance(startDate, swapTenor_,
					ModifiedFollowing,
					endOfMonth_);
			}
			else {
				endDate = startDate + swapTenor_;
			}
		}

		Schedule baseSchedule(startDate, endDate,
			baseLegTenor_,
			baseLegCalendar_,
			baseLegConvention_,
			baseLegTerminationDateConvention_,
			baseLegRule_,
			endOfMonth_);

		Schedule basisSchedule(startDate, endDate,
			basisLegTenor_,
			basisLegCalendar_,
			basisLegConvention_,
			basisLegTerminationDateConvention_,
			basisLegRule_,
			endOfMonth_);

		Rate usedbasisSpread = spread_;
		if (spread_ == Null<Rate>()) {
			QL_REQUIRE(!basisLegIndex_->forwardingTermStructure().empty(),
				"null term structure set to this instance of " <<
				basisLegIndex_->name());
			GenericSwap temp(type_, nominal_,
				baseSchedule,
				baseSchedule,
				baseLegIndex_,
				baseLegDayCount_,
				basisSchedule,
				basisSchedule,
				basisLegIndex_,
				basisLegDayCount_,
				0.0);
			// ATM on the forecasting curve
			bool includeSettlementDateFlows = false;
			temp.setPricingEngine(boost::shared_ptr<PricingEngine>(
				new DiscountingSwapEngine(
				basisLegIndex_->forwardingTermStructure(),
				includeSettlementDateFlows)));
			usedbasisSpread = temp.fairSpread();
		}

		boost::shared_ptr<GenericSwap> iborbasis(new
			GenericSwap(type_, nominal_,
			baseSchedule,
			baseSchedule,
			baseLegIndex_,
			baseLegDayCount_,
			basisSchedule,
			basisSchedule,
			basisLegIndex_,
			basisLegDayCount_,
			0.0,
			usedbasisSpread));
		iborbasis->setPricingEngine(engine_);
		return iborbasis;
	}


	MakeGenericSwap& MakeGenericSwap::withType(GenericSwap::Type type) {
		type_ = type;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withNominal(Real n) {
		nominal_ = n;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withSettlementDays(Natural fixingDays) {
		fixingDays_ = fixingDays;
		effectiveDate_ = Date();
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withEffectiveDate(const Date& effectiveDate) {
		effectiveDate_ = effectiveDate;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withTerminationDate(const Date& terminationDate) {
		terminationDate_ = terminationDate;
		swapTenor_ = Period();
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withPaymentConvention(BusinessDayConvention bdc) {
		paymentConvention_ = bdc;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withEndOfMonth(bool flag) {
		endOfMonth_ = flag;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBaseLegTenor(const Period& t) {
		baseLegTenor_ = t;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBaseLegCalendar(const Calendar& cal) {
		baseLegCalendar_ = cal;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBaseLegConvention(BusinessDayConvention bdc) {
		baseLegConvention_ = bdc;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBaseLegTerminationDateConvention(
		BusinessDayConvention bdc) {
		baseLegTerminationDateConvention_ = bdc;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBaseLegRule(DateGeneration::Rule r) {
		baseLegRule_ = r;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBaseLegDayCount(const DayCounter& dc) {
		baseLegDayCount_ = dc;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBasisLegTenor(const Period& t) {
		basisLegTenor_ = t;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBasisLegCalendar(const Calendar& cal) {
		basisLegCalendar_ = cal;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBasisLegConvention(BusinessDayConvention bdc) {
		basisLegConvention_ = bdc;
		return *this;
	}
	MakeGenericSwap& MakeGenericSwap::withBasisLegTerminationDateConvention(
		BusinessDayConvention bdc) {
		basisLegTerminationDateConvention_ = bdc;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBasisLegRule(DateGeneration::Rule r) {
		basisLegRule_ = r;
		if (r == DateGeneration::Zero)
			basisLegTenor_ = Period(Once);
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBasisLegDayCount(const DayCounter& dc) {
		basisLegDayCount_ = dc;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withBasisLegSpread(Spread sp) {
		spread_ = sp;
		return *this;
	}

	MakeGenericSwap& MakeGenericSwap::withDiscountingTermStructure(
		const Handle<YieldTermStructure>& discountingTermStructure) {
		engine_ = boost::shared_ptr<PricingEngine>(new
			DiscountingSwapEngine(discountingTermStructure));
		return *this;
	}
}

