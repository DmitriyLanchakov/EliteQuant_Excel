#include <math/distributions/gammadistribution.hpp>
#include <boost/math/special_functions/gamma.hpp>

namespace QLExtension {

	InverseCumulativeGamma::InverseCumulativeGamma(Real shape, Real scale)
		: dist_(boost::math::gamma_distribution<>(shape, scale))
	{}
		

	Real InverseCumulativeGamma::operator()(Real x) const {
		return boost::math::quantile(dist_, x);
	}
}
