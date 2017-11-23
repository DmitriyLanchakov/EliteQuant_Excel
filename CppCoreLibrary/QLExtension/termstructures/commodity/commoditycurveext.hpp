#ifndef qlex_commodity_curve_ext_hpp
#define qlex_commodity_curve_ext_hpp

#include <ql/termstructure.hpp>
#include <ql/currency.hpp>
#include <ql/math/interpolations/forwardflatinterpolation.hpp>
#include <ql/math/interpolations/linearinterpolation.hpp>
#include <ql/math/interpolations/loginterpolation.hpp>
#include <ql/time/calendars/nullcalendar.hpp>
#include <ql/time/daycounters/actual365fixed.hpp>

using namespace QuantLib;

namespace QLExtension {

    //! Commodity term structure
	class CommodityCurveExt : public QuantLib::TermStructure {
        friend class CommodityIndexExt;
      public:
        // constructor
		  CommodityCurveExt(const std::string name,		// e.g. commod ng exchange
                       const std::vector<Date>& dates,   // expiry date
                       const std::vector<Real>& prices,		// quoted price
					   const Calendar& calendar = QuantLib::NullCalendar(),
                       const DayCounter& dayCounter = Actual365Fixed());

		  CommodityCurveExt(const std::string name,
					   const Calendar& calendar = QuantLib::NullCalendar(),
                       const DayCounter& dayCounter = Actual365Fixed());

        //! \name Inspectors
        //@{
        const std::string name() const;
        Date maxDate() const;
        const std::vector<Time>& times() const;
        const std::vector<Date>& dates() const;
        const std::vector<Real>& prices() const;
        std::vector<std::pair<Date,Real> > nodes() const;		// date/price pair
        bool empty() const;

        void setPrices(std::map<Date, Real>& prices);

		// forward price on the curve
        Real price(const Date& d) const;

        friend std::ostream& operator<<(std::ostream& out,
                                        const CommodityCurveExt& curve);
      protected:
		std::string name_;
        mutable std::vector<Date> dates_;
        mutable std::vector<Time> times_;
        mutable std::vector<Real> data_;
        mutable Interpolation interpolation_;
		LogLinear interpolator_;

        Real priceImpl(Time t) const;
    };


    // inline definitions

	inline bool operator==(const CommodityCurveExt& c1, const CommodityCurveExt& c2) {
        return c1.name() == c2.name();
    }

	inline const std::string CommodityCurveExt::name() const {
        return name_;
    }

	inline Date CommodityCurveExt::maxDate() const {
        return dates_.back();
    }

	inline const std::vector<Time>& CommodityCurveExt::times() const {
        return times_;
    }

	inline const std::vector<Date>& CommodityCurveExt::dates() const {
        return dates_;
    }

	inline const std::vector<Real>& CommodityCurveExt::prices() const {
        return data_;
    }

	inline bool CommodityCurveExt::empty() const {
        return dates_.empty();
    }

	inline std::vector<std::pair<Date, Real> > CommodityCurveExt::nodes() const {
        std::vector<std::pair<Date,Real> > results(dates_.size());
        for (Size i = 0; i < dates_.size(); ++i)
            results[i] = std::make_pair(dates_[i], data_[i]);
        return results;
    }

    // gets a price that can include an arbitrary number of basis curves
	inline Real CommodityCurveExt::price(const Date& d) const {
        Date date = d;
        Time t = timeFromReference(date);
        Real priceValue = 0;
        try {
            priceValue = priceImpl(t);
        } catch (const std::exception& e) {
            QL_FAIL("error retrieving price for curve [" << name() << "]: "
                    << e.what());
        }
        return priceValue;
    }

	inline Real CommodityCurveExt::priceImpl(Time t) const {
        return interpolation_(t, true);		// allowExtrapolation = true
    }
}


#endif
