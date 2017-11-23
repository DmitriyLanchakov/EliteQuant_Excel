#include <instruments/energycommodityext.hpp>
#include <iomanip>

namespace QLExtension {
    EnergyDailyPositionExt::EnergyDailyPositionExt()
    : payLegPrice(0), receiveLegPrice(0), unrealized(false) {}

    EnergyDailyPositionExt::EnergyDailyPositionExt(const Date& date,
                                             Real payLegPrice,
                                             Real receiveLegPrice,
                                             bool unrealized)
    : date(date), quantityAmount(0), payLegPrice(payLegPrice),
      receiveLegPrice(receiveLegPrice), unrealized(unrealized) {}

    std::ostream& operator<<(std::ostream& out,
                             const EnergyDailyPositionExts& dailyPositions) {
        out << std::setw(12) << std::left << "positions"
            << std::setw(12) << std::right << "pay"
            << std::setw(12) << std::right << "receive"
            << std::setw(10) << std::right << "qty"
            << std::setw(14) << std::right << "delta"
            << std::setw(10) << std::right << "open" << std::endl;

        for (EnergyDailyPositionExts::const_iterator i = dailyPositions.begin();
             i != dailyPositions.end(); ++i) {
            const EnergyDailyPositionExt& dailyPosition = i->second;
            out << std::setw(4) << io::iso_date(i->first) << "  "
                << std::setw(12) << std::right << std::fixed
                << std::setprecision(6) << dailyPosition.payLegPrice
                << std::setw(12) << std::right << std::fixed
                << std::setprecision(6) << dailyPosition.receiveLegPrice
                << std::setw(10) << std::right << std::fixed
                << std::setprecision(2) << dailyPosition.quantityAmount
                << std::setw(14) << std::right << std::fixed
                << std::setprecision(2) << dailyPosition.riskDelta
                << std::setw(10) << std::right << std::fixed
                << std::setprecision(2)
                << (dailyPosition.unrealized ? dailyPosition.quantityAmount : 0)
                << std::endl;
        }

        return out;
    }


    void EnergyCommodityExt::setupArguments(PricingEngine::arguments* args) const {
        EnergyCommodityExt::arguments* arguments =
            dynamic_cast<EnergyCommodityExt::arguments*>(args);
        QL_REQUIRE(arguments != 0, "wrong argument type");
        //arguments->legs = legs_;
        //arguments->payer = payer_;
    }

    void EnergyCommodityExt::fetchResults(const PricingEngine::results* r) const {
        Instrument::fetchResults(r);
        const EnergyCommodityExt::results* results =
            dynamic_cast<const EnergyCommodityExt::results*>(r);
        QL_REQUIRE(results != 0, "wrong result type");
    }

    EnergyCommodityExt::EnergyCommodityExt(const std::string& commodityName)
    : commodityName_(commodityName) {}

    std::string EnergyCommodityExt::commodityName()  {
		return commodityName_;
    }
}

