#ifndef qlex_energy_swap_ext_hpp
#define qlex_energy_swap_ext_hpp

#include <instruments/energycommodityext.hpp>
#include <instruments/pricingperiodext.hpp>
#include <cashflows/commoditycashflow.hpp>
#include <ql/time/calendar.hpp>

using namespace QuantLib;

namespace QLExtension {

    class EnergySwapExt : public EnergyCommodityExt {
      public:
        EnergySwapExt(const PricingPeriodExts& pricingPeriods,
                   const std::string& commodityName,
				   const Frequency deliverySchedule,
				   const Calendar& calendar);

        bool isExpired() const;
        const Calendar& calendar() const { return calendar_; }
        const PricingPeriodExts& pricingPeriods() const { return pricingPeriods_; }
        const EnergyDailyPositionExts& dailyPositions() const {
            return dailyPositions_;
        }

        Real quantity() const;

      protected:
        Calendar calendar_;
		Frequency deliverySchedule_;		// daily or monthly
        PricingPeriodExts pricingPeriods_;
		// daily delivery throughout swap periods
		// e.g. 6m swap 30 per month, dailyPositions has 6*30 days with 1 per day
        mutable EnergyDailyPositionExts dailyPositions_;
    };

}

#endif
