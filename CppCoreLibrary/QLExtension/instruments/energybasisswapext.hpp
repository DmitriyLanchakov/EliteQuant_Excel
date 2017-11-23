#ifndef qlex_energy_basis_swap_ext_hpp
#define qlex_energy_basis_swap_ext_hpp

#include <instruments/energyswapext.hpp>
#include <indexes/commodityindexext.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>

using namespace QuantLib;

namespace QLExtension {

    //! Energy basis swap
    class EnergyBasisSwapExt : public EnergySwapExt {
      public:
        EnergyBasisSwapExt(
                    const boost::shared_ptr<CommodityIndexExt>& payIndex,
					const boost::shared_ptr<CommodityIndexExt>& receiveIndex,
                    const PricingPeriodExts& pricingPeriods,			// shared by pay/rec legs
                    const std::string& commodityNamee,
                    const Handle<YieldTermStructure>& payLegTermStructure,
                    const Handle<YieldTermStructure>& receiveLegTermStructure,
					const Frequency deliverySchedule = Frequency::Daily,
					const Calendar& calendar = QuantLib::NullCalendar());

		const boost::shared_ptr<CommodityIndexExt>& payIndex() const {
            return payIndex_;
        }
		const boost::shared_ptr<CommodityIndexExt>& receiveIndex() const {
            return receiveIndex_;
        }

      protected:
        void performCalculations() const;

		boost::shared_ptr<CommodityIndexExt> payIndex_;
		boost::shared_ptr<CommodityIndexExt> receiveIndex_;
        Handle<YieldTermStructure> payLegTermStructure_;
        Handle<YieldTermStructure> receiveLegTermStructure_;
    };

}

#endif
