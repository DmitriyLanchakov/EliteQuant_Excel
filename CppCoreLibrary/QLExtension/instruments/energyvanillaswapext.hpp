#ifndef qlex_energy_vanilla_swap_ext_hpp
#define qlex_energy_vanilla_swap_ext_hpp

#include <instruments/energyswapext.hpp>
#include <indexes/commodityindexext.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>

using namespace QuantLib;

namespace QLExtension {

    //! Vanilla energy swap, S-K
	// NGH5 matures at end of Feb, and delivers every day amount = totalamoutn/31; 
	//		the payoff happens at payment date, or end of Feb
    class EnergyVanillaSwapExt : public EnergySwapExt {
      public:
        EnergyVanillaSwapExt(
                    bool payer,
                    const Real fixedPrice,
					const boost::shared_ptr<CommodityIndexExt>& index,
                    const PricingPeriodExts& pricingPeriods,
                    const std::string& commodityName,
                    const Handle<YieldTermStructure>& payLegTermStructure,
                    const Handle<YieldTermStructure>& receiveLegTermStructure,
					const Frequency deliverySchedule = Frequency::Daily,
					const Calendar& calendar = QuantLib::NullCalendar());

        bool isExpired() const;
        Integer payReceive() const { return payReceive_; }
        const Real fixedPrice() const { return fixedPrice_; }
		const boost::shared_ptr<CommodityIndexExt>& index() const {
            return index_;
        }

      protected:
        void performCalculations() const;

        Integer payReceive_;
        Real fixedPrice_;
		boost::shared_ptr<CommodityIndexExt> index_;
        Handle<YieldTermStructure> payLegTermStructure_;
        Handle<YieldTermStructure> receiveLegTermStructure_;
    };

}


#endif
