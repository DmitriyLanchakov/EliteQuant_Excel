#include <termstructures/commodity/commoditycurveext.hpp>

namespace QLExtension {

	CommodityCurveExt::CommodityCurveExt(const std::string name,
                                   const std::vector<Date>& dates,
                                   const std::vector<Real>& prices,
								   const Calendar& calendar,
                                   const DayCounter& dayCounter)
    : TermStructure(0, calendar, dayCounter),
      name_(name), dates_(dates), data_(prices),
	  interpolator_(LogLinear()) {

        QL_REQUIRE(dates_.size()>1, "too few dates");
        QL_REQUIRE(data_.size()==dates_.size(), "dates/prices count mismatch");

		// it assumes dates[0] has times[0] = 0
        times_.resize(dates_.size());
        times_[0]=0.0;
        for (Size i = 1; i < dates_.size(); i++) {
            QL_REQUIRE(dates_[i] > dates_[i-1],
                       "invalid date (" << dates_[i] << ", vs "
                       << dates_[i-1] << ")");
            times_[i] = dayCounter.yearFraction(dates_[0], dates_[i]);
        }

        interpolation_ =
            interpolator_.interpolate(times_.begin(), times_.end(),
                                      data_.begin());
        interpolation_.update();
    }

	CommodityCurveExt::CommodityCurveExt(const std::string name,
                                   const Calendar& calendar,
                                   const DayCounter& dayCounter)
    : TermStructure(0, calendar, dayCounter),
	name_(name), interpolator_(LogLinear()) {}

	void CommodityCurveExt::setPrices(std::map<Date, Real>& prices) {
        QL_REQUIRE(prices.size()>1, "too few prices");

        dates_.clear();
        data_.clear();
        for (std::map<Date, Real>::const_iterator i = prices.begin(); i != prices.end(); ++i) {
            dates_.push_back(i->first);
            data_.push_back(i->second);
        }

        times_.resize(dates_.size());
        times_[0]=0.0;
        for (Size i = 1; i < dates_.size(); i++)
            times_[i] = dayCounter().yearFraction(dates_[0], dates_[i]);

        interpolation_ =
            interpolator_.interpolate(times_.begin(), times_.end(),
                                      data_.begin());
        interpolation_.update();
    }

	std::ostream& operator<<(std::ostream& out, const CommodityCurveExt& curve) {
        out << "[" << curve.name_ << "]";
        return out;
    }

}
