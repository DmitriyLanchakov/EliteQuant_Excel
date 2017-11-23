#ifndef qlex_energy_commodity_ext_hpp
#define qlex_energy_commodity_ext_hpp

#include <ql/instrument.hpp>
#include <ql/money.hpp>
#include <ql/time/date.hpp>
#include <vector>
#include <iosfwd>

using namespace QuantLib;

namespace QLExtension {

    struct EnergyDailyPositionExt {
        Date date;
        Real quantityAmount;
        Real payLegPrice;
        Real receiveLegPrice;
        Real riskDelta;
        bool unrealized;

        EnergyDailyPositionExt();
        EnergyDailyPositionExt(const Date& date,
                            Real payLegPrice,
                            Real receiveLegPrice,
                            bool unrealized);
    };

    typedef std::map<Date, EnergyDailyPositionExt> EnergyDailyPositionExts;

    #ifndef __DOXYGEN__
    std::ostream& operator<<(std::ostream& out,
                             const EnergyDailyPositionExts& dailyPositions);
    #endif



    //! Energy commodity class
    /*! \ingroup instruments */
    class EnergyCommodityExt : public Instrument {
      public:
        class arguments;
        class results;
        class engine;

        enum DeliverySchedule { Constant,
                                Window,
                                Hourly,
                                Daily,
                                Weekly,
                                Monthly,
                                Quarterly,
                                Yearly };
        enum QuantityPeriodicity { Absolute,
                                   PerHour,
                                   PerDay,
                                   PerWeek,
                                   PerMonth,
                                   PerQuarter,
                                   PerYear };
        enum PaymentSchedule { WindowSettlement,
                               MonthlySettlement,
                               QuarterlySettlement,
                               YearlySettlement };

        EnergyCommodityExt(const std::string& commodityName);

        virtual Real quantity() const = 0;
        std::string commodityName();

        void setupArguments(PricingEngine::arguments*) const;
        void fetchResults(const PricingEngine::results*) const;

      protected:
        std::string commodityName_;			// name of this trade
    };


    class EnergyCommodityExt::arguments : public virtual PricingEngine::arguments {
      public:
        void validate() const {}
    };

    class EnergyCommodityExt::results : public Instrument::results {
      public:
        Real NPV;
        void reset() {
            Instrument::results::reset();
        }
    };

    class EnergyCommodityExt::engine
        : public GenericEngine<EnergyCommodityExt::arguments,
                               EnergyCommodityExt::results> {};

}

#endif
