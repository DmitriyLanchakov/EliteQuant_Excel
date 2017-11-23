// (1) schedule is n+1. in FloatingLeg <ql/cashflows/cashflowvectors.hpp> and #include <ql/utilities/vectors.hpp>
// It takes index min(i,n), where n = schedule.size()-1
// that is, schedule has start date

// (2) composition. Eventually it needs to separate payment schedule from fixing schedule.
// but a quick fix is to add composition factor, e.g. 1mth/3mth would be 3, and a high-frequency schedule
// In class IborLeg, add composition to IborCoupon and overload indexFixing() to composition

#ifndef qlex_ibor_basis_swap_hpp
#define qlex_ibor_basis_swap_hpp

#include <ql/instruments/swap.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/time/daycounter.hpp>
#include <ql/time/schedule.hpp>
#include <boost/optional.hpp>

using namespace QuantLib;

namespace QLExtension {

	class QuantLib::IborIndex;
	class QuantLib::OvernightIndex;

	//! Ibor basis swap, e.g. 1m vs 3m
	class GenericSwap : public Swap {
	public:
		enum Type { Receiver = -1, Payer = 1 };		// rec/pay first leg
		// fixed rate with notional and notionals
		GenericSwap(
			Type type,
			Real nominal,
			const Schedule& firstLegResetSchedule,			// fixed leg
			const Schedule& firstLegPaymenttSchedule,			// fixed leg
			const Rate firstLegRate,					// fixed rate
			const DayCounter& firstLegDayCount,
			const Schedule& secondLegResetSchedule,
			const Schedule& secondLegPaymentSchedule,
			const boost::shared_ptr<IborIndex>& secondLegIndex,
			const DayCounter& secondLegDayCount,
			Spread secondLegSpread = 0.0,
			bool arithmeticAveragedCoupon = false,
			boost::optional<BusinessDayConvention> firstLegConvention =
			boost::none,
			boost::optional<BusinessDayConvention> secondLegConvention =
			boost::none);
		GenericSwap(
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
			Spread secondLegSpread = 0.0,
			bool arithmeticAveragedCoupon = false,
			boost::optional<BusinessDayConvention> firstLegConvention =
			boost::none,
			boost::optional<BusinessDayConvention> secondLegConvention =
			boost::none);

		// basis leg with notional and notionals
		GenericSwap(
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
			Spread firstLegSpread = 0.0,
			Spread secondLegSpread = 0.0,
			bool arithmeticAveragedCoupon = false,
			boost::optional<BusinessDayConvention> firstLegConvention =
			boost::none,
			boost::optional<BusinessDayConvention> secondLegConvention =
			boost::none);
		GenericSwap(
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
			Spread firstLegSpread = 0.0,
			Spread secondLegSpread = 0.0,
			bool arithmeticAveragedCoupon = false,
			boost::optional<BusinessDayConvention> firstLegConvention =
			boost::none,
			boost::optional<BusinessDayConvention> secondLegConvention =
			boost::none);
		//! \name Inspectors
		//@{
		Type type() const { return type_; }
		Real nominal() const;
		
		std::vector<Real> firstLegNominals() const { return firstLegNominals_; }
		const Schedule& firstLegResetSchedule() const { return firstLegResetSchedule_; }
		const Schedule& firstLegPaymentSchedule() const { return firstLegPaymentSchedule_; }
		const boost::shared_ptr<IborIndex>& firstLegIndex() const { return firstLegIndex_; }
		const DayCounter& firstLegDayCount() const { return firstLegDayCount_; }
		const BusinessDayConvention& firstLegConvention() const { return firstLegConvention_; }

		std::vector<Real> secondLegNominals() const { return secondLegNominals_; }
		const Schedule& secondLegResetSchedule() { return secondLegResetSchedule_; }
		const Schedule& secondLegPaymentSchedule() { return secondLegPaymentSchedule_; }
		const boost::shared_ptr<IborIndex>& secondLegIndex() { return secondLegIndex_; }
		const DayCounter& secondLegDayCount() const { return secondLegDayCount_; }
		const BusinessDayConvention& secondLegConvention() const { return secondLegConvention_; }

		Rate fixedRate() { return firstLegRate_; }
		Spread firstLegSpread() { return firstLegSpread_; }
		Spread secondLegSpread() { return secondLegSpread_; }
		const Leg& firstLeg() const { return legs_[0]; }
		const Leg& secondLeg() const { return legs_[1]; }
		const std::vector<std::string>& firstLegInfo() const { return firstLegInfo_; }
		const std::vector<std::string>& secondLegInfo() const { return secondLegInfo_; }
		//@}

		//! \name Results
		//@{
		Real firstLegBPS() const;
		Real firstLegNPV() const;
		Real fairRate() const;

		Real secondLegBPS() const;
		Real secondLegNPV() const;
		Spread fairSpread() const;
		//@}
	private:
		void initialize(bool isfixed=false);
		Type type_;
		bool arithmeticAveragedCoupon_;

		std::vector<Real> firstLegNominals_;
		Schedule firstLegResetSchedule_;
		Schedule firstLegPaymentSchedule_;
		boost::shared_ptr<IborIndex> firstLegIndex_;
		DayCounter firstLegDayCount_;
		BusinessDayConvention firstLegConvention_;
		Rate firstLegRate_;	
		Spread firstLegSpread_;

		std::vector<Real> secondLegNominals_;
		Schedule secondLegResetSchedule_;
		Schedule secondLegPaymentSchedule_;
		boost::shared_ptr<IborIndex> secondLegIndex_;
		boost::shared_ptr<OvernightIndex> overnightIndex_;
		DayCounter secondLegDayCount_;
		BusinessDayConvention secondLegConvention_;
		Spread secondLegSpread_;

		std::vector<std::string> firstLegInfo_;
		std::vector<std::string> secondLegInfo_;
	};

	// inline
	inline Real GenericSwap::nominal() const {
		QL_REQUIRE(firstLegNominals_.size() == 1, "varying nominals");
		return firstLegNominals_[0];
	}
}

#endif
