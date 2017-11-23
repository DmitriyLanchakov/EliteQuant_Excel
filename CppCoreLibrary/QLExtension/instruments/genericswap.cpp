#include <ql/cashflows/overnightindexedcoupon.hpp>
#include <ql/cashflows/iborcoupon.hpp>
#include <ql/cashflows/fixedratecoupon.hpp>
#include <cashflows/overnightindexedcoupon.hpp>
#include <instruments/genericswap.hpp>
//#include <iomanip>

namespace QLExtension {
	GenericSwap::GenericSwap(
		Type type,
		Real nominal,
		const Schedule& firstLegResetSchedule,			// fixed leg
		const Schedule& firstLegPaymenttSchedule,
		const Rate firstLegRate,					// fixed rate
		const DayCounter& firstLegDayCount,
		const Schedule& secondLegResetSchedule,
		const Schedule& secondLegPaymentSchedule,
		const boost::shared_ptr<IborIndex>& secondLegIndex,
		const DayCounter& secondLegDayCount,
		Spread secondLegSpread,
		bool arithmeticAveragedCoupon,
		boost::optional<BusinessDayConvention> firstLegConvention,
		boost::optional<BusinessDayConvention> secondLegConvention)
		: type_(type), Swap(2), firstLegNominals_(std::vector<Real>(1, nominal)), secondLegNominals_(std::vector<Real>(1, nominal)),
		firstLegResetSchedule_(firstLegResetSchedule), firstLegPaymentSchedule_(firstLegPaymenttSchedule), firstLegRate_(firstLegRate), firstLegDayCount_(firstLegDayCount),
		secondLegResetSchedule_(secondLegResetSchedule), secondLegPaymentSchedule_(secondLegPaymentSchedule), secondLegIndex_(secondLegIndex), secondLegDayCount_(secondLegDayCount),
		firstLegSpread_(0.0), secondLegSpread_(secondLegSpread), arithmeticAveragedCoupon_(arithmeticAveragedCoupon){

		if (firstLegConvention)
			firstLegConvention_ = *firstLegConvention;
		else
			firstLegConvention_ = firstLegPaymentSchedule_.businessDayConvention();

		if (secondLegConvention)
			secondLegConvention_ = *secondLegConvention;
		else
			secondLegConvention_ = secondLegPaymentSchedule_.businessDayConvention();

		initialize(true);
	}

	GenericSwap::GenericSwap(
		Type type,
		std::vector<Real> firstLegNominals,
		const Schedule& firstLegResetSchedule,			// fixed leg
		const Schedule& firstLegPaymenttSchedule,
		const Rate firstLegRate,					// fixed rate
		const DayCounter& firstLegDayCount,
		std::vector<Real> secoondLegNominals,
		const Schedule& secondLegResetSchedule,
		const Schedule& secondLegPaymentSchedule,
		const boost::shared_ptr<IborIndex>& secondLegIndex,
		const DayCounter& secondLegDayCount,
		Spread secondLegSpread,
		bool arithmeticAveragedCoupon,
		boost::optional<BusinessDayConvention> firstLegConvention,
		boost::optional<BusinessDayConvention> secondLegConvention)
		: type_(type), Swap(2), firstLegNominals_(firstLegNominals), secondLegNominals_(secoondLegNominals),
		firstLegResetSchedule_(firstLegResetSchedule), firstLegPaymentSchedule_(firstLegPaymenttSchedule), firstLegRate_(firstLegRate), firstLegDayCount_(firstLegDayCount),
		secondLegResetSchedule_(secondLegResetSchedule), secondLegPaymentSchedule_(secondLegPaymentSchedule), secondLegIndex_(secondLegIndex), secondLegDayCount_(secondLegDayCount),
		firstLegSpread_(0.0), secondLegSpread_(secondLegSpread), arithmeticAveragedCoupon_(arithmeticAveragedCoupon){

		if (firstLegConvention)
			firstLegConvention_ = *firstLegConvention;
		else
			firstLegConvention_ = firstLegPaymentSchedule_.businessDayConvention();

		if (secondLegConvention)
			secondLegConvention_ = *secondLegConvention;
		else
			secondLegConvention_ = secondLegPaymentSchedule.businessDayConvention();

		initialize(true);
	}

	GenericSwap::GenericSwap(
		Type type,
		Real nominal,
		const Schedule& firstLegResetSchedule,			// LIB3M
		const Schedule& firstLegPaymenttSchedule,
		const boost::shared_ptr<IborIndex>& firstLegIndex,
		const DayCounter& firstLegDayCount,
		const Schedule& secondLegResetSchedule,
		const Schedule& secondLegPaymentSchedule,
		const boost::shared_ptr<IborIndex>& secondLegIndex,
		const DayCounter& secondLegDayCount,
		Spread firstLegSpread,
		Spread secondLegSpread,
		bool arithmeticAveragedCoupon,
		boost::optional<BusinessDayConvention> firstLegConvention,
		boost::optional<BusinessDayConvention> secondLegConvention)
		: type_(type), Swap(2), firstLegNominals_(std::vector<Real>(1, nominal)), secondLegNominals_(std::vector<Real>(1, nominal)),
		firstLegResetSchedule_(firstLegResetSchedule), firstLegPaymentSchedule_(firstLegPaymenttSchedule), firstLegIndex_(firstLegIndex), firstLegDayCount_(firstLegDayCount),
		secondLegResetSchedule_(secondLegResetSchedule), secondLegPaymentSchedule_(secondLegPaymentSchedule), secondLegIndex_(secondLegIndex), secondLegDayCount_(secondLegDayCount),
		firstLegSpread_(firstLegSpread), secondLegSpread_(secondLegSpread), arithmeticAveragedCoupon_(arithmeticAveragedCoupon){

		if (firstLegConvention)
			firstLegConvention_ = *firstLegConvention;
		else
			firstLegConvention_ = firstLegPaymentSchedule_.businessDayConvention();

		if (secondLegConvention)
			secondLegConvention_ = *secondLegConvention;
		else
			secondLegConvention_ = secondLegPaymentSchedule_.businessDayConvention();

		initialize(false);
	}

	GenericSwap::GenericSwap(
		Type type,
		std::vector<Real> firstLegNominals,
		const Schedule& firstLegResetSchedule,			// LIB3M
		const Schedule& firstLegPaymenttSchedule,
		const boost::shared_ptr<IborIndex>& firstLegIndex,
		const DayCounter& firstLegDayCount,
		std::vector<Real> secondLegNominals,
		const Schedule& secondLegResetSchedule,
		const Schedule& secondLegPaymentSchedule,
		const boost::shared_ptr<IborIndex>& secondLegIndex,
		const DayCounter& secondLegDayCount,
		Spread firstLegSpread,
		Spread secondLegSpread,
		bool arithmeticAveragedCoupon,
		boost::optional<BusinessDayConvention> firstLegConvention,
		boost::optional<BusinessDayConvention> secondLegConvention)
		: type_(type), Swap(2), firstLegNominals_(firstLegNominals), secondLegNominals_(secondLegNominals),
		firstLegResetSchedule_(firstLegResetSchedule), firstLegPaymentSchedule_(firstLegPaymenttSchedule), firstLegIndex_(firstLegIndex), firstLegDayCount_(firstLegDayCount),
		secondLegResetSchedule_(secondLegResetSchedule), secondLegPaymentSchedule_(secondLegPaymentSchedule), secondLegIndex_(secondLegIndex), secondLegDayCount_(secondLegDayCount),
		firstLegSpread_(firstLegSpread), secondLegSpread_(secondLegSpread), arithmeticAveragedCoupon_(arithmeticAveragedCoupon){

		if (firstLegConvention)
			firstLegConvention_ = *firstLegConvention;
		else
			firstLegConvention_ = firstLegPaymentSchedule_.businessDayConvention();

		if (secondLegConvention)
			secondLegConvention_ = *secondLegConvention;
		else
			secondLegConvention_ = secondLegPaymentSchedule_.businessDayConvention();

		initialize(false);
	}

	void GenericSwap::initialize(bool isfixed) {
		// (FIXED, LIB), (FIXED, OIS), (LIB, OIS), (LIB, LIB)
		// Leg 0, either Fixed or LIB3M
		if (isfixed)
		{
			// ignore firstLegResetSchedule for now
			legs_[0] = FixedRateLeg(firstLegPaymentSchedule_)
				.withNotionals(firstLegNominals_)
				.withCouponRates(firstLegRate_, firstLegDayCount_)
				.withPaymentAdjustment(firstLegConvention_);
		}
		else
		{
			legs_[0] = IborLeg(firstLegPaymentSchedule_, firstLegIndex_)
				.withNotionals(firstLegNominals_)
				.withPaymentDayCounter(firstLegDayCount_)
				.withPaymentAdjustment(firstLegConvention_)
				.withSpreads(firstLegSpread_);
		} // leg 0 ends

		// Leg 1, either OIS or LIB1M
		if (arithmeticAveragedCoupon_)     // ois
		{
			overnightIndex_ =
				boost::dynamic_pointer_cast<OvernightIndex> (secondLegIndex_);
			legs_[1] = OvernightLeg(secondLegPaymentSchedule_, overnightIndex_)
				.withNotionals(secondLegNominals_)
				.withPaymentDayCounter(secondLegDayCount_)
				.withPaymentAdjustment(secondLegConvention_)
				.withSpreads(secondLegSpread_);
		}
		else
		{
			legs_[1] = IborLeg(secondLegPaymentSchedule_, secondLegIndex_)
				.withNotionals(secondLegNominals_)
				.withPaymentDayCounter(secondLegDayCount_)
				.withPaymentAdjustment(secondLegConvention_)
				.withSpreads(secondLegSpread_);
		}

		for (Size j = 0; j<2; ++j) {
			for (Leg::iterator i = legs_[j].begin(); i != legs_[j].end(); ++i)
				registerWith(*i);
		}

		switch (type_) {
		case Payer:
			payer_[0] = -1.0;
			payer_[1] = +1.0;
			break;
		case Receiver:
			payer_[0] = +1.0;
			payer_[1] = -1.0;
			break;
		default:
			QL_FAIL("Unknown generic swap type");
		}

		// intialize leg info
		// have to downcast from c++ because c# can't see derived type
		for (int lg = 0; lg < 2; lg++)
		{
			std::string startdate = "", enddate = "", paymentdate = "", resetdate = "";
			double balance = 0, rate = 0, spread = 0, payment = 0, discount = 0, pv = 0;

			int n = legs_[lg].size();
			for (int i = 0; i < n; i++)
			{
				boost::shared_ptr<FixedRateCoupon> cf1 = boost::dynamic_pointer_cast<FixedRateCoupon>(legs_[lg][i]);
				boost::shared_ptr<FloatingRateCoupon> cf2;
				if (cf1)			// fixed leg
				{
					startdate = std::to_string(cf1->accrualStartDate().serialNumber());
					enddate = std::to_string(cf1->accrualEndDate().serialNumber());
					paymentdate = enddate;
					resetdate = "";
					balance = cf1->nominal();
					rate = cf1->rate();
					spread = 0.0;
					payment = cf1->amount();
				}
				else       // (!cf1) floating leg
				{
					cf2 = boost::dynamic_pointer_cast<FloatingRateCoupon>(legs_[lg][i]);

					startdate = std::to_string(cf2->accrualStartDate().serialNumber());
					enddate = std::to_string(cf2->accrualEndDate().serialNumber());
					paymentdate = enddate;
					resetdate = std::to_string(cf2->fixingDate().serialNumber());
					balance = cf2->nominal();
					rate = cf2->rate();
					spread = cf2->spread();
					payment = cf2->amount();
				}

				std::ostringstream os;
				//os << std::setprecision(2) << startdate << "," << enddate << "," << paymentdate << "," << resetdate << "," << balance << "," << rate << "," << spread << "," << payment;
				os  << startdate << "," << enddate << "," << paymentdate << "," << resetdate << "," << balance << "," << rate << "," << spread << "," << payment;

				if (lg == 0)
				{
					firstLegInfo_.push_back(os.str());
				}
				else
				{
					secondLegInfo_.push_back(os.str());
				}
			}
		}	
	}

	Real GenericSwap::fairRate() const {
		static Spread basisPoint = 1.0e-4;
		calculate();
		return firstLegRate_ - NPV_ / (firstLegBPS() / basisPoint);
	}

	Spread GenericSwap::fairSpread() const {
		static Spread basisPoint = 1.0e-4;
		calculate();
		return secondLegSpread_ - NPV_ / (secondLegBPS() / basisPoint);
	}

	Real GenericSwap::firstLegBPS() const {
		calculate();
		QL_REQUIRE(legBPS_[0] != Null<Real>(), "result not available");
		return legBPS_[0];
	}

	Real GenericSwap::secondLegBPS() const {
		calculate();
		QL_REQUIRE(legBPS_[1] != Null<Real>(), "result not available");
		return legBPS_[1];
	}

	Real GenericSwap::firstLegNPV() const {
		calculate();
		QL_REQUIRE(legNPV_[0] != Null<Real>(), "result not available");
		return legNPV_[0];
	}

	Real GenericSwap::secondLegNPV() const {
		calculate();
		QL_REQUIRE(legNPV_[1] != Null<Real>(), "result not available");
		return legNPV_[1];
	}
}

