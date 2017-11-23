#include <cashflows/cashflowvectorsext.hpp>
#include <ql/cashflows/fixedratecoupon.hpp>
#include <ql/cashflows/capflooredcoupon.hpp>
#include <ql/cashflows/rangeaccrual.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/time/schedule.hpp>

namespace AMP {

    namespace detail {

        Rate effectiveFixedRate(const std::vector<Spread>& spreads,
                                const std::vector<Rate>& caps,
                                const std::vector<Rate>& floors,
                                Size i) {
            Rate result = QuantLib::detail::get(spreads, i, 0.0);
            Rate floor = QuantLib::detail::get(floors, i, Null<Rate>());
            if (floor!=Null<Rate>())
                result = std::max(floor, result);
			Rate cap = QuantLib::detail::get(caps, i, Null<Rate>());
            if (cap!=Null<Rate>())
                result = std::min(cap, result);
            return result;
        }

        bool noOption(const std::vector<Rate>& caps,
                      const std::vector<Rate>& floors,
                      Size i) {
			return (QuantLib::detail::get(caps, i, Null<Rate>()) == Null<Rate>()) &&
				(QuantLib::detail::get(floors, i, Null<Rate>()) == Null<Rate>());
        }

    }

}
