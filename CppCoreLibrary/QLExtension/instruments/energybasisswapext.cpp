#include <instruments/energybasisswapext.hpp>

namespace QLExtension {

    EnergyBasisSwapExt::EnergyBasisSwapExt(
                    const boost::shared_ptr<CommodityIndexExt>& payIndex,
					const boost::shared_ptr<CommodityIndexExt>& receiveIndex,
                    const PricingPeriodExts& pricingPeriods,
                    const std::string& commodityName,
                    const Handle<YieldTermStructure>& payLegTermStructure,
                    const Handle<YieldTermStructure>& receiveLegTermStructure,
					const Frequency deliverySchedule,
					const Calendar& calendar)
    : EnergySwapExt(pricingPeriods, commodityName, deliverySchedule, calendar),
	  payIndex_(payIndex), receiveIndex_(receiveIndex), 
	  payLegTermStructure_(payLegTermStructure),
	  receiveLegTermStructure_(receiveLegTermStructure) {
        QL_REQUIRE(pricingPeriods_.size() > 0, "no payment dates");
		registerWith(Settings::instance().evaluationDate());
        registerWith(payIndex_);
        registerWith(receiveIndex_);
    }

	void EnergyBasisSwapExt::performCalculations() const {

		try {

			if (payIndex_->empty()) {
				if (payIndex_->forwardCurveEmpty()) {
					QL_FAIL("index [" + payIndex_->name() +
						"] does not have any quotes or forward prices");
				}
			}
			if (receiveIndex_->empty()) {
				if (receiveIndex_->forwardCurveEmpty()) {
					QL_FAIL("index [" + receiveIndex_->name() +
						"] does not have any quotes or forward prices");
				}
			}

			NPV_ = 0.0;
			additionalResults_.clear();
			dailyPositions_.clear();

			Date evaluationDate = Settings::instance().evaluationDate();

			Date lastPayIndexQuoteDate = payIndex_->lastQuoteDate();
			Date lastReceiveIndexQuoteDate = receiveIndex_->lastQuoteDate();

			if (lastPayIndexQuoteDate < evaluationDate - 1) {
				std::ostringstream message;
				message << "index [" << payIndex_->name()
					<< "] has last quote date of "
					<< io::iso_date(lastPayIndexQuoteDate);
			}
			if (lastReceiveIndexQuoteDate < evaluationDate - 1) {
				std::ostringstream message;
				message << "index [" << receiveIndex_->name()
					<< "] has last quote date of "
					<< io::iso_date(lastReceiveIndexQuoteDate);
			}

			Date lastQuoteDate = std::min(lastPayIndexQuoteDate,
				lastReceiveIndexQuoteDate);

			Real totalQuantityAmount = 0;

			// price each period
			for (PricingPeriodExts::const_iterator pi = pricingPeriods_.begin();
				pi != pricingPeriods_.end(); ++pi) {
				const boost::shared_ptr<PricingPeriodExt>& pricingPeriod = *pi;

				// Add: skip expired periods
				if (pricingPeriod->paymentDate() <= evaluationDate)
					continue;

				Integer periodDayCount = 0;

				// get the index quotes
				Date periodStartDate =
					calendar_.adjust(pricingPeriod->startDate());
				for (Date stepDate = periodStartDate;
					stepDate <= pricingPeriod->endDate();
					stepDate = calendar_.advance(stepDate, 1 * Days)) {

					bool unrealized = stepDate > evaluationDate;
					Real payQuoteValue = 0;
					Real receiveQuoteValue = 0;

					if (stepDate <= lastQuoteDate) {
						payQuoteValue = payIndex_->price(stepDate);
						receiveQuoteValue = receiveIndex_->price(stepDate);
					}
					else {
						payQuoteValue = payIndex_->forwardPrice(stepDate);
						receiveQuoteValue =	receiveIndex_->forwardPrice(stepDate);
					}

					if (payQuoteValue == 0) {
						std::ostringstream message;
						message << "pay quote value for curve ["
							<< payIndex_->name() << "] is 0 for date "
							<< io::iso_date(stepDate);
					}
					if (receiveQuoteValue == 0) {
						std::ostringstream message;
						message << "receive quote value for curve ["
							<< receiveIndex_->name() << "] is 0 for date "
							<< io::iso_date(stepDate);
					}

					QL_REQUIRE(payQuoteValue != Null<Real>(),
						"curve [" << payIndex_->name() <<
						"] missing value for pricing date: "
						<< stepDate);
					QL_REQUIRE(receiveQuoteValue != Null<Real>(),
						"curve [" << receiveIndex_->name() <<
						"] missing value for pricing date: "
						<< stepDate);

					Real payLegPriceValue = payQuoteValue;
					Real receiveLegPriceValue =	receiveQuoteValue;

					payLegPriceValue = pricingPeriod->payCoeff() * payLegPriceValue + pricingPeriod->paySpread();
					receiveLegPriceValue = pricingPeriod->recCoeff() * receiveLegPriceValue + pricingPeriod->recSpread();

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
				if (pricingPeriod->paymentDate() >= evaluationDate + 2 /* settlement days*/) {
					payLegDiscountFactor =
						payLegTermStructure_->discount(
						pricingPeriod->paymentDate());
					receiveLegDiscountFactor =
						receiveLegTermStructure_->discount(
						pricingPeriod->paymentDate());
				}

				Real uDelta = receiveLegValue + payLegValue;
				Real dDelta = (receiveLegValue * receiveLegDiscountFactor) +
					(payLegValue * payLegDiscountFactor);
				Real pmtDiscountFactor =
					(dDelta  > 0) ? receiveLegDiscountFactor : payLegDiscountFactor;

				pricingPeriod->setuPayDelta(payLegValue);
				pricingPeriod->setuRecDelta(receiveLegValue);
				pricingPeriod->setdPayDelta(payLegValue * payLegDiscountFactor);
				pricingPeriod->setdRecDelta(receiveLegValue * receiveLegDiscountFactor);

				pricingPeriod->setFinalized(pricingPeriod->paymentDate() <= evaluationDate);

				NPV_ += dDelta;
			}

			additionalResults_["dailyPositions"] = dailyPositions_;
		}
		catch (const QuantLib::Error&) {
			throw;
		}
		catch (const std::exception&) {
			throw;
		}
    }
}