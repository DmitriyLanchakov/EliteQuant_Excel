#include <instruments/energyfutureext.hpp>
#include <ql/settings.hpp>

namespace QLExtension {

    EnergyFutureExt::EnergyFutureExt(
                      Integer buySell, const boost::shared_ptr<PricingPeriodExt>& pricingPeriod,
                      const Real tradePrice,
					  const boost::shared_ptr<CommodityIndexExt>& index,
                      const std::string& commodityName,
					  const Handle<YieldTermStructure>& discountTermStructure)
    : EnergyCommodityExt(commodityName),
      buySell_(buySell), pricingPeriod_(pricingPeriod), tradePrice_(tradePrice),
	  index_(index), discountTermStructure_(discountTermStructure) {
        registerWith(Settings::instance().evaluationDate());
        registerWith(index_);
    }

    bool EnergyFutureExt::isExpired() const {
        return false;
    }

    void EnergyFutureExt::performCalculations() const {

        NPV_ = 0.0;
        additionalResults_.clear();

        Date evaluationDate = Settings::instance().evaluationDate();
		Date paymentDate = pricingPeriod_->paymentDate();
		Date lastQuoteDate = index_->lastQuoteDate();

        Real quoteValue = 0;

		if (paymentDate <= lastQuoteDate) {
			quoteValue = index_->price(paymentDate);
		}
		else {
			quoteValue = index_->forwardPrice(paymentDate);
		}

        QL_REQUIRE(quoteValue != Null<Real>(),
                   "missing quote for [" << index_->name() << "]");

		Real tradePriceValue = tradePrice_;
		Real quotePriceValue = quoteValue;
		Real quantityAmount = pricingPeriod_->quantity();
		Real discountFactor = discountTermStructure_->discount(pricingPeriod_->paymentDate());

		// no discount
        //Real delta = (((quotePriceValue - tradePriceValue) * quantityAmount)
        //              * index_->lotQuantity()) * buySell_ * discountFactor;

		Real payLegValue, receiveLegValue;
		if (buySell_ == 1)
		{
			payLegValue = -tradePriceValue*quantityAmount*index_->lotQuantity();
			receiveLegValue = quotePriceValue*quantityAmount*index_->lotQuantity();
		}
		else
		{
			receiveLegValue = -tradePriceValue*quantityAmount*index_->lotQuantity();
			payLegValue = quotePriceValue*quantityAmount*index_->lotQuantity();
		}

		Real delta = (payLegValue + receiveLegValue) * discountFactor;

		pricingPeriod_->setuPayDelta(payLegValue);
		pricingPeriod_->setuRecDelta(receiveLegValue);
		pricingPeriod_->setdPayDelta(payLegValue * discountFactor);
		pricingPeriod_->setdRecDelta(receiveLegValue * discountFactor);

		pricingPeriod_->setFinalized(pricingPeriod_->paymentDate() <= evaluationDate);

		NPV_ = delta * ((pricingPeriod_->paymentDate() <= evaluationDate) ? 0 : 1.0);

        // additionalResults_["brokerCommission"] =
        //     -(brokerCommissionValue * quantityAmount);
    }

}

