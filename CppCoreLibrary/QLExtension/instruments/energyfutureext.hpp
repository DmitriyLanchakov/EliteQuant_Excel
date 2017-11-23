#ifndef qlex_energy_future_ext_hpp
#define qlex_energy_future_ext_hpp

#include <instruments/energycommodityext.hpp>
#include <indexes/commodityindexext.hpp>
#include <instruments/pricingperiodext.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>

using namespace QuantLib;

namespace QLExtension {

    //! Energy future
    /*! \ingroup instruments */
    class EnergyFutureExt : public EnergyCommodityExt {
      public:
        EnergyFutureExt(Integer buySell, const boost::shared_ptr<PricingPeriodExt>& pricingPeriod,
                     const Real tradePrice,				// USD3.5 quantity = 20, unit of measure = 1000MMBTU --> set quantity = 20000MMBTU
					 const boost::shared_ptr<CommodityIndexExt>& index,
                     const std::string& commodityName,
					 const Handle<YieldTermStructure>& discountTermStructure);
        bool isExpired() const;
        //Integer buySell{} const { return buySell_; }
		boost::shared_ptr<PricingPeriodExt> pricingPeriod() const { return pricingPeriod_; }
        const Real tradePrice() const { return tradePrice_; }
		const boost::shared_ptr<CommodityIndexExt> index() const { return index_; }

		Real quantity() const {
			return pricingPeriod_->quantity();
		}

      protected:
        void performCalculations() const;
        Integer buySell_;	// payer = 1, receiver = -1
		boost::shared_ptr<PricingPeriodExt> pricingPeriod_;
        Real tradePrice_;
		boost::shared_ptr<CommodityIndexExt> index_;
		Handle<YieldTermStructure> discountTermStructure_;
    };
}

#endif
