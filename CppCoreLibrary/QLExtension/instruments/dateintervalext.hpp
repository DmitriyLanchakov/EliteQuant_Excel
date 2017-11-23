#ifndef qlex_date_interval_ext_hpp
#define qlex_date_interval_ext_hpp

#include <ql/time/date.hpp>
#include <ql/errors.hpp>

using namespace QuantLib;

namespace QLExtension {

    //! Date interval described by a number of a given time unit
    /*! \ingroup datetime */
    class DateIntervalExt {
        friend std::ostream& operator<<(std::ostream&, const DateIntervalExt&);

      private:
        Date startDate_;
        Date endDate_;
      public:
        DateIntervalExt()  {}
        DateIntervalExt(const Date& startDate, const Date& endDate)
        : startDate_(startDate), endDate_(endDate) {
            QL_REQUIRE(endDate_ >= startDate_,
                       "end date must be >= start date");
        }
        const Date& startDate() const { return startDate_; }
        const Date& endDate() const { return endDate_; }

        bool isDateBetween(Date date,
                           bool includeFirst = true,
                           bool includeLast = true) const {
            if (includeFirst && !(date >= startDate_))
                return false;
            else if (!(date > startDate_))
                return false;
            if (includeLast && !(date <= endDate_))
                return false;
            else if (!(date < endDate_))
                return false;
            return true;
        }

        DateIntervalExt intersection(const DateIntervalExt& di) const {
            if ((startDate_ < di.startDate_ && endDate_ < di.startDate_) ||
                (startDate_ > di.endDate_ && endDate_ > di.endDate_))
                return DateIntervalExt();
            return DateIntervalExt(std::max(startDate_, di.startDate_),
                                std::min(endDate_, di.endDate_));
        }

        bool operator==(const DateIntervalExt& rhs) const {
            return startDate_ == rhs.startDate_ && endDate_ == rhs.endDate_;
        }

        bool operator!=(const DateIntervalExt& rhs) const {
            return !(*this == rhs);
        }
    };

}

#endif
