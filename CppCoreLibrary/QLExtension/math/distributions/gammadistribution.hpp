// http://www.alglib.net/specialfunctions/incompletegamma.php
// http://www.boost.org/doc/libs/1_35_0/libs/math/doc/sf_and_dist/html/math_toolkit/special/sf_gamma/igamma_inv.html
// http://www.cs.washington.edu/research/projects/uns/F9/src/boost-1.39.0/libs/math/doc/sf_and_dist/html/math_toolkit/special/sf_gamma/igamma.html
// http://en.wikipedia.org/wiki/Gamma_distribution
// http://en.wikipedia.org/wiki/Incomplete_gamma_function
#ifndef qlex_math_gamma_distribution_h
#define qlex_math_gamma_distribution_h

#include <ql/errors.hpp>
#include <ql/types.hpp>
#include <functional>
#include <boost/math/distributions/gamma.hpp>
#include <boost/math/special_functions/gamma.hpp>

using namespace QuantLib;

namespace QLExtension {
	class InverseCumulativeGamma
		: public std::unary_function < Real, Real > {
	public:
		InverseCumulativeGamma(Real shape = 1.0,
			Real scale = 1.0);

		// in boost p is lower incomplete gamma function, q is upper incomplete gamma function
		// gamma_p(a,z) = p = gamma(a,z)/Gamma(a)
		Real operator()(Real x) const;
		
		boost::math::gamma_distribution<> dist_;
	};
}


#endif
