/*! \file svi.hpp
    \brief SVI functions
*/

#ifndef qlex_svi_hpp
#define qlex_svi_hpp

#include <ql/types.hpp>
using namespace QuantLib;

namespace QLExtension {

    Real unsafeSviVolatility(Rate strike,
                              Rate forward,
                              Time expiryTime,
							  Real a,
							  Real b,
							  Real rho,
							  Real m,
							  Real sigma);

    Real sviVolatility(Rate strike,
                        Rate forward,
                        Time expiryTime,
						Real a,
						Real b,
						Real rho,
						Real m,
						Real sigma);

    void validateSviParameters(Real a,
                                Real b,
								Real rho,
								Real m,
								Real sigma);

}

#endif
