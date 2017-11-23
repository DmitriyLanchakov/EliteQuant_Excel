#include <instruments/energyswapext.hpp>
#include <ql/settings.hpp>

namespace QLExtension {

    EnergySwapExt::EnergySwapExt(
                      const PricingPeriodExts& pricingPeriods,
                      const std::string& commodityName,
					  const Frequency deliverySchedule,
					  const Calendar& calendar)
    : EnergyCommodityExt(commodityName), deliverySchedule_(deliverySchedule),
      calendar_(calendar), pricingPeriods_(pricingPeriods) {}

    Real EnergySwapExt::quantity() const {
        Real totalQuantityAmount = 0;
		for (PricingPeriodExts::const_iterator pi = pricingPeriods_.begin();
             pi != pricingPeriods_.end(); ++pi) {
            totalQuantityAmount += (*pi)->quantity();
        }
        return totalQuantityAmount;
    }

    bool EnergySwapExt::isExpired() const {
        return pricingPeriods_.empty()
            || detail::simple_event(pricingPeriods_.back()->paymentDate())
               .hasOccurred();
    }

}

