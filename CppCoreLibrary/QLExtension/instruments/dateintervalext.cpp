#include <instruments/dateintervalext.hpp>

namespace QLExtension {

    std::ostream& operator<<(std::ostream& out, const DateIntervalExt& di) {
        if (di.startDate_ == Date() || di.endDate_ == Date())
            return out << "Null<DateIntervalExt>()";
        return out << di.startDate_ << " to " << di.endDate_;
    }

}
