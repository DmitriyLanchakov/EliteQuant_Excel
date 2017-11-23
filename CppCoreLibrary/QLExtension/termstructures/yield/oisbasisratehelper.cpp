#include <termstructures/yield/oisbasisratehelper.hpp>
#include <instruments/makeiboroisbasisswap.hpp>
#include <instruments/makegenericswap.hpp>
#include <ql/instruments/makeois.hpp>
#include <ql/instruments/makevanillaswap.hpp>
#include <ql/pricingengines/swap/discountingswapengine.hpp>
#include <ql/cashflows/overnightindexedcoupon.hpp>

using boost::shared_ptr;

namespace QLExtension {

	namespace {
		void no_deletion(YieldTermStructure*) {}
	}

	IBOROISBasisRateHelper::IBOROISBasisRateHelper(
		Natural settlementDays,
		const Period& tenor,
		const Handle<Quote>& overnightSpread,
		const boost::shared_ptr<IborIndex>& iborIndex,
		const boost::shared_ptr<OvernightIndex>& overnightIndex,
		const Handle<YieldTermStructure>& discount)
		: RelativeDateRateHelper(overnightSpread),
		settlementDays_(settlementDays), tenor_(tenor),
		iborIndex_(iborIndex), overnightIndex_(overnightIndex),
		discountHandle_(discount) {
		registerWith(iborIndex_);
		registerWith(overnightIndex_);
		registerWith(discountHandle_);
		initializeDates();
	}

	void IBOROISBasisRateHelper::initializeDates() {

		// dummy OvernightIndex with curve/swap arguments
		// review here
		boost::shared_ptr<IborIndex> tempClonedIborIndex =
			overnightIndex_->clone(termStructureHandle_);
		shared_ptr<OvernightIndex> clonedOvernightIndex =
			boost::dynamic_pointer_cast<OvernightIndex>(tempClonedIborIndex);

		// input discount curve Handle might be empty now but it could
		//    be assigned a curve later; use a RelinkableHandle here
		swap_ = MakeIBOROISBasisSwap(tenor_, iborIndex_, clonedOvernightIndex, 0.0)
			.withDiscountingTermStructure(discountRelinkableHandle_)
			.withSettlementDays(settlementDays_);

		earliestDate_ = swap_->startDate();
		latestDate_ = swap_->maturityDate();
	}

	void IBOROISBasisRateHelper::setTermStructure(YieldTermStructure* t) {
		// do not set the relinkable handle as an observer -
		// force recalculation when needed
		bool observer = false;

		shared_ptr<YieldTermStructure> temp(t, no_deletion);
		termStructureHandle_.linkTo(temp, observer);

		if (discountHandle_.empty())
			discountRelinkableHandle_.linkTo(temp, observer);
		else
			discountRelinkableHandle_.linkTo(*discountHandle_, observer);

		RelativeDateRateHelper::setTermStructure(t);
	}

	Real IBOROISBasisRateHelper::impliedQuote() const {
		QL_REQUIRE(termStructure_ != 0, "term structure not set");
		// we didn't register as observers - force calculation
		swap_->recalculate();
		return swap_->fairSpread();
	}

	void IBOROISBasisRateHelper::accept(AcyclicVisitor& v) {
		Visitor<IBOROISBasisRateHelper>* v1 =
			dynamic_cast<Visitor<IBOROISBasisRateHelper>*>(&v);
		if (v1 != 0)
			v1->visit(*this);
		else
			RateHelper::accept(v);
	}


	FixedOISBasisRateHelper::FixedOISBasisRateHelper(
		Natural settlementDays,
		const Period& tenor, // swap maturity
		const Handle<Quote>& overnightSpread,
		const Handle<Quote>& fixedRate,
		Frequency fixedFrequency,
		BusinessDayConvention fixedConvention,
		const DayCounter& fixedDayCount,
		const boost::shared_ptr<OvernightIndex>& overnightIndex,
		Frequency overnightFrequency,
		const Handle<YieldTermStructure>& discount)
		: RelativeDateRateHelper(overnightSpread), fixedRate_(fixedRate),
		usedFixedRate_(fixedRate->value()), fixedFrequency_(fixedFrequency),
		fixedConvention_(fixedConvention), fixedDayCount_(fixedDayCount),
		settlementDays_(settlementDays), tenor_(tenor),
		overnightIndex_(overnightIndex), overnightFrequency_(overnightFrequency),
		discountHandle_(discount) {
		registerWith(fixedRate);
		registerWith(overnightIndex_);
		registerWith(discountHandle_);
		initializeDates();
	}

	void FixedOISBasisRateHelper::initializeDates() {

		// dummy OvernightIndex with curve/swap arguments
		// review here
		boost::shared_ptr<IborIndex> clonedIborIndex =
			overnightIndex_->clone(termStructureHandle_);
		shared_ptr<OvernightIndex> clonedOvernightIndex =
			boost::dynamic_pointer_cast<OvernightIndex>(clonedIborIndex);

		// input discount curve Handle might be empty now but it could
		//    be assigned a curve later; use a RelinkableHandle here
		boost::shared_ptr<IborIndex> dummyIndex(new IborIndex("Dummy",
			Period(overnightFrequency_),
			settlementDays_,
			clonedIborIndex->currency(),
			clonedIborIndex->fixingCalendar(),
			clonedIborIndex->businessDayConvention(),
			clonedIborIndex->endOfMonth(),
			clonedIborIndex->dayCounter()));
		boost::shared_ptr<VanillaSwap> dummyVanillaSwap =
			MakeVanillaSwap(tenor_, dummyIndex, usedFixedRate_)
			.withFixedLegDayCount(fixedDayCount_)
			.withFixedLegTenor(Period(fixedFrequency_))
			.withFixedLegConvention(fixedConvention_)
			.withFixedLegTerminationDateConvention(fixedConvention_);
		boost::shared_ptr<OvernightIndexedSwap> dummyOISSwap =
			MakeOIS(tenor_, clonedOvernightIndex, usedFixedRate_)
			.withSettlementDays(settlementDays_)
			.withPaymentFrequency(overnightFrequency_);
		Leg oisBasisLeg = dummyOISSwap->overnightLeg();

		boost::shared_ptr<FloatingRateCouponPricer> arithmeticPricer(
			new ArithmeticAveragedOvernightIndexedCouponPricer());
		for (Size i = 0; i < oisBasisLeg.size(); i++) {
			boost::shared_ptr<OvernightIndexedCoupon> c =
				boost::dynamic_pointer_cast<OvernightIndexedCoupon> (oisBasisLeg[i]);
			c->setPricer(arithmeticPricer);
		}

		swap_ = boost::shared_ptr<Swap>(new Swap(dummyVanillaSwap->fixedLeg(), dummyOISSwap->overnightLeg()));
		boost::shared_ptr<PricingEngine> engine(new DiscountingSwapEngine(discountRelinkableHandle_));
		swap_->setPricingEngine(engine);

		earliestDate_ = swap_->startDate();
		latestDate_ = swap_->maturityDate();
	}

	void FixedOISBasisRateHelper::setTermStructure(YieldTermStructure* t) {
		// do not set the relinkable handle as an observer -
		// force recalculation when needed
		bool observer = false;

		shared_ptr<YieldTermStructure> temp(t, no_deletion);
		termStructureHandle_.linkTo(temp, observer);

		if (discountHandle_.empty())
			discountRelinkableHandle_.linkTo(temp, observer);
		else
			discountRelinkableHandle_.linkTo(*discountHandle_, observer);

		RelativeDateRateHelper::setTermStructure(t);
	}

	Real FixedOISBasisRateHelper::impliedQuote() const {
		QL_REQUIRE(termStructure_ != 0, "term structure not set");
		// we didn't register as observers - force calculation
		swap_->recalculate();
		return -swap_->NPV() / (swap_->legBPS(1) / 1.0e-4);
	}

	void FixedOISBasisRateHelper::accept(AcyclicVisitor& v) {
		Visitor<FixedOISBasisRateHelper>* v1 =
			dynamic_cast<Visitor<FixedOISBasisRateHelper>*>(&v);
		if (v1 != 0)
			v1->visit(*this);
		else
			RateHelper::accept(v);
	}

	void FixedOISBasisRateHelper::update() {
		if (usedFixedRate_ != fixedRate_->value()) {
			usedFixedRate_ = fixedRate_->value();
			initializeDates();
		}
		RelativeDateRateHelper::update();
	}

	IBORBasisRateHelper::IBORBasisRateHelper(
		Natural settlementDays,
		const Period& tenor, // swap maturity
		const Handle<Quote>& basis,
		const boost::shared_ptr<IborIndex>& baseLegIborIndex,
		const boost::shared_ptr<IborIndex>& basisLegIborIndex,
		const Handle<YieldTermStructure>& discountingCurve)
		: RelativeDateRateHelper(basis),
		settlementDays_(settlementDays), tenor_(tenor),
		baseLegIndex_(baseLegIborIndex),
		discountHandle_(discountingCurve) {

		basisLegIndex_ = basisLegIborIndex->clone(termStructureHandle_);
		registerWith(baseLegIndex_);
		registerWith(basisLegIndex_);
		registerWith(discountHandle_);
		initializeDates();
	}

	void IBORBasisRateHelper::initializeDates() {

		// input discount curve Handle might be empty now but it could
		//    be assigned a curve later; use a RelinkableHandle here
		swap_ = MakeGenericSwap(tenor_, baseLegIndex_, basisLegIndex_, 0.0)
			.withDiscountingTermStructure(discountRelinkableHandle_)
			.withSettlementDays(settlementDays_);
		boost::shared_ptr<PricingEngine> engine(new DiscountingSwapEngine(discountRelinkableHandle_));
		swap_->setPricingEngine(engine);

		earliestDate_ = swap_->startDate();
		latestDate_ = swap_->maturityDate();
	}

	void IBORBasisRateHelper::setTermStructure(YieldTermStructure* t) {
		// do not set the relinkable handle as an observer -
		// force recalculation when needed
		bool observer = false;

		// set basis index forwarding termstructure
		shared_ptr<YieldTermStructure> temp(t, no_deletion);
		termStructureHandle_.linkTo(temp, observer);

		if (discountHandle_.empty())
			discountRelinkableHandle_.linkTo(temp, observer);
		else
			discountRelinkableHandle_.linkTo(*discountHandle_, observer);

		RelativeDateRateHelper::setTermStructure(t);
	}

	Real IBORBasisRateHelper::impliedQuote() const {
		QL_REQUIRE(termStructure_ != 0, "term structure not set");
		// we didn't register as observers - force calculation
		swap_->recalculate();
		return swap_->fairSpread();
	}

	void IBORBasisRateHelper::accept(AcyclicVisitor& v) {
		Visitor<IBORBasisRateHelper>* v1 =
			dynamic_cast<Visitor<IBORBasisRateHelper>*>(&v);
		if (v1 != 0)
			v1->visit(*this);
		else
			RateHelper::accept(v);
	}

	void IBORBasisRateHelper::update() {
		RelativeDateRateHelper::update();
	}
}