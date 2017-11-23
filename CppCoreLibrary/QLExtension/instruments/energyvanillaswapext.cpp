#include <instruments/energyvanillaswapext.hpp>

namespace QLExtension {

	EnergyVanillaSwapExt::EnergyVanillaSwapExt(
                    bool payer,
                    const Real fixedPrice,
					const boost::shared_ptr<CommodityIndexExt>& index,
                    const PricingPeriodExts& pricingPeriods,
                    const std::string& commodityName,
                    const Handle<YieldTermStructure>& payLegTermStructure,
                    const Handle<YieldTermStructure>& receiveLegTermStructure,
					const Frequency deliverySchedule,
					const Calendar& calendar)
    : EnergySwapExt(pricingPeriods, commodityName, deliverySchedule, calendar),
      payReceive_(payer ? 1 : 0), fixedPrice_(fixedPrice),
      index_(index), payLegTermStructure_(payLegTermStructure),
      receiveLegTermStructure_(receiveLegTermStructure) {

        QL_REQUIRE(pricingPeriods_.size() > 0, "no pricing periods");
		registerWith(Settings::instance().evaluationDate());
        registerWith(index_);
    }

	bool EnergyVanillaSwapExt::isExpired() const {
        return detail::simple_event(pricingPeriods_.back()->endDate())
               .hasOccurred();
    }

	void EnergyVanillaSwapExt::performCalculations() const {

        try {
			// no fixings, no forwarding curve
            if (index_->empty()) {
                if (index_->forwardCurveEmpty()) {
                    QL_FAIL("index [" << index_->name()
                            << "] does not have any quotes");
                }
            }

            NPV_ = 0.0;
            additionalResults_.clear();
            dailyPositions_.clear();

            Date evaluationDate = Settings::instance().evaluationDate();

            Date lastQuoteDate = index_->lastQuoteDate();
            if (lastQuoteDate < evaluationDate - 1) {
                std::ostringstream message;
                message << "index [" << index_->name()
                        << "] has last quote date of "
                        << io::iso_date(lastQuoteDate);
            }

            Real totalQuantityAmount = 0;

            // price each period
            for (PricingPeriodExts::const_iterator pi = pricingPeriods_.begin();
                 pi != pricingPeriods_.end(); ++pi) {
                const boost::shared_ptr<PricingPeriodExt>& pricingPeriod = *pi;

				QL_REQUIRE(pricingPeriod->quantity() != 0,
                           "quantity is zero");

				// Add: skip expired periods
				if (pricingPeriod->paymentDate() <= evaluationDate)
					continue;

                Integer periodDayCount = 0;

                // get the futures quotes or everything after
				// for everyday in the period
                Date periodStartDate =
					calendar_.adjust(pricingPeriod->startDate());
                for (Date stepDate = periodStartDate;
                     stepDate <= pricingPeriod->endDate();
                     stepDate = calendar_.advance(stepDate, 1*Days)) {

                    bool unrealized = stepDate > evaluationDate;
                    Real quoteValue = 0;

                    if (stepDate <= lastQuoteDate) {
                        quoteValue = index_->price(stepDate);
                    } else {
                        quoteValue = index_->forwardPrice(stepDate);
                    }

                    if (quoteValue == 0) {
                        std::ostringstream message;
                        message << "pay quote value for curve ["
                                << index_->name() << "] is 0 for date "
                                << io::iso_date(stepDate);
                    }

                    QL_REQUIRE(quoteValue != Null<Real>(),
                               "curve [" << index_->name() <<
                               "] missing value for pricing date: "
                               << stepDate);

                    Real fixedLegPriceValue = fixedPrice_;
					Real floatingLegPriceValue = quoteValue;
                    Real payLegPriceValue =
                        payReceive_ > 0 ? fixedLegPriceValue :
                                          floatingLegPriceValue;
                    Real receiveLegPriceValue =
                        payReceive_ > 0 ? floatingLegPriceValue :
                                          fixedLegPriceValue;

                    dailyPositions_[stepDate] =
                        EnergyDailyPositionExt(stepDate, payLegPriceValue,
                                            receiveLegPriceValue, unrealized);
                    periodDayCount++;

					// only consider daily and monthly
					// if monthly, assume periodStartDate is the beginning of the month
					// and price this day only
					if (deliverySchedule_ == Frequency::Monthly)
					{
						break;
					}
                }

				Real periodQuantityAmount = pricingPeriod->quantity();
                totalQuantityAmount += periodQuantityAmount;

                Real avgDailyQuantityAmount =
                    periodDayCount == 0 ? 0 :
                                          periodQuantityAmount / periodDayCount;

				// how much to pay at payment date, based on (daily) delivery positions
                Real payLegValue = 0;
                Real receiveLegValue = 0;
                for (std::map<Date, EnergyDailyPositionExt>::iterator dpi =
                         dailyPositions_.find(periodStartDate);
                     dpi != dailyPositions_.end() &&
                         dpi->first <= pricingPeriod->endDate(); ++dpi) {
                    EnergyDailyPositionExt& dailyPosition = dpi->second;
                    dailyPosition.quantityAmount = avgDailyQuantityAmount;
                    dailyPosition.riskDelta =
                        (-dailyPosition.payLegPrice + dailyPosition.receiveLegPrice) * avgDailyQuantityAmount;
                    payLegValue += -dailyPosition.payLegPrice * avgDailyQuantityAmount;
                    receiveLegValue += dailyPosition.receiveLegPrice * avgDailyQuantityAmount;
                }

                Real payLegDiscountFactor = 1;
                Real receiveLegDiscountFactor = 1;
				if (pricingPeriod->paymentDate() >= evaluationDate + 2) {
                    payLegDiscountFactor =
                        payLegTermStructure_->discount(
                                                pricingPeriod->paymentDate());
                    receiveLegDiscountFactor =
                        receiveLegTermStructure_->discount(
                                                pricingPeriod->paymentDate());
                }

                Real uDelta = receiveLegValue + payLegValue;		// undiscounted
                Real dDelta = (receiveLegValue * receiveLegDiscountFactor) +
                    (payLegValue * payLegDiscountFactor);			// discounted

				Real pmtDiscountFactor =
                    (dDelta  > 0) ? receiveLegDiscountFactor :
                                    payLegDiscountFactor;

				pricingPeriod->setuPayDelta(payLegValue);
				pricingPeriod->setuRecDelta(receiveLegValue);
				pricingPeriod->setdPayDelta(payLegValue * payLegDiscountFactor);
				pricingPeriod->setdRecDelta(receiveLegValue * receiveLegDiscountFactor);
				
				pricingPeriod->setFinalized(pricingPeriod->paymentDate() <= evaluationDate);

                NPV_ += dDelta;
            }	// end for each period

            additionalResults_["dailyPositions"] = dailyPositions_;

        } catch (const QuantLib::Error&) {
            throw;
        } catch (const std::exception&) {
            throw;
        }
    }

}

