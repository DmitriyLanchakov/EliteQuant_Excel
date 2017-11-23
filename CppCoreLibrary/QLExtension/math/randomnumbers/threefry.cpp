#include <boost/array.hpp>
#include <boost/static_assert.hpp>
#include <boost/cstdint.hpp>
#include <boost/limits.hpp>
#include <boost/mpl/range_c.hpp>
#include <boost/mpl/for_each.hpp>
#include <math/randomnumbers/detail/rotl.hpp>
#include <math/randomnumbers/threefry.hpp>

namespace boost{
namespace random{

const unsigned
threefry_constants<2, uint32_t>::Rotations[]  =
    {13, 15, 26, 6, 17, 29, 16, 24};

const unsigned
threefry_constants<4, uint32_t>::Rotations0[]  = 
    {10, 11, 13, 23, 6, 17, 25, 18};

const unsigned
threefry_constants<4, uint32_t>::Rotations1[]  = 
    {26, 21, 27, 5, 20, 11, 10, 20};

const unsigned
threefry_constants<2, uint64_t>::Rotations[]  =
    {16, 42, 12, 31, 16, 32, 24, 21};

const unsigned
threefry_constants<4, uint64_t>::Rotations0[]  = 
    {14, 52, 23, 5, 25, 46, 58, 32};

const unsigned
threefry_constants<4, uint64_t>::Rotations1[]  = {
    16, 57, 40, 37, 33, 12, 22, 32};


} // namespace random
} // namespace boost
