#ifndef amp_commodity_hpp
#define amp_commodity_hpp

#include <ql/instrument.hpp>
#include <ql/money.hpp>
#include <vector>
#include <iosfwd>

using namespace QuantLib;

namespace AMP {

    typedef std::map<std::string, boost::any> SecondaryCosts;		// additional cost
    typedef std::map<std::string, Money> SecondaryCostAmounts;		// additional cost amount

    std::ostream& operator<<(std::ostream& out,
                             const SecondaryCostAmounts& secondaryCostAmounts);

	struct PricingError {
		enum Level { Info, Warning, Error, Fatal };

		Level errorLevel;
		std::string tradeId;
		std::string error;
		std::string detail;

		PricingError(Level errorLevel,
			const std::string& error,
			const std::string& detail)
			: errorLevel(errorLevel), error(error), detail(detail) {}
	};

	typedef std::vector<PricingError> PricingErrors;
    std::ostream& operator<<(std::ostream& out, const PricingError& error);		// log error message
    std::ostream& operator<<(std::ostream& out, const PricingErrors& errors);


    //! Commodity base class
    /*! \ingroup instruments */
    class Commodity : public Instrument {
      public:
        Commodity(const boost::shared_ptr<SecondaryCosts>& secondaryCosts);
        const boost::shared_ptr<SecondaryCosts>& secondaryCosts() const;
        const SecondaryCostAmounts& secondaryCostAmounts() const;
        const PricingErrors& pricingErrors() const;
        void addPricingError(PricingError::Level errorLevel,
                             const std::string& error,
                             const std::string& detail = "") const;
      protected:
        boost::shared_ptr<SecondaryCosts> secondaryCosts_;
        mutable PricingErrors pricingErrors_;
        mutable SecondaryCostAmounts secondaryCostAmounts_;
    };

}

#endif
