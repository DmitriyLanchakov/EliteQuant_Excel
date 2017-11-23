#include <ql/types.hpp>
#include <ql/math/randomnumbers/seedgenerator.hpp>
#include <ql/math/randomnumbers/rngtraits.hpp>
#include <ql/math/distributions/normaldistribution.hpp>
#include <ql/methods/montecarlo/sample.hpp>
#include <boost/static_assert.hpp>
#include <boost/random.hpp>

#include <math/randomnumbers/counter_based_engine.hpp>
#include <math/randomnumbers/threefry.hpp>
#include <math/randomnumbers/philox.hpp>
#include <math/randomnumbers/boostrngbinding.hpp>
#include <settings.hpp>

using namespace QuantLib;

namespace QLExtension {
	// Returns uniform rsg, caller responsible for array size
	void originalthreefry(uint64_t seed, uint64_t skip, uint64_t buffer[], uint64_t length)
	{
		boost::random::counter_based_engine<uint64_t, boost::random::threefry<2, uint64_t>, 32> rng(seed);
		rng.restart({0,0});		// fixed it
		rng.discard(skip);

		for (uint64_t i = 0; i < length; i++)
		{
			buffer[i] = rng();
		}
	}

	// Returns uniform rsg, caller responsible for array size
	void uniformthreefry(uint64_t seed, uint64_t skip, double buffer[], uint64_t length)
	{
		BoostThreefryUniformRng rng(seed);
		rng.restart(0);		// fixed it
		rng.discard(skip);
		
		for (uint64_t i = 0; i < length; i++)
		{
			buffer[i] = rng.nextReal();
		}
	}

	// Returns normal rsg, caller responsible for array size
	void normalthreefry(uint64_t seed, uint64_t skip, double buffer[], uint64_t length)
	{
		InverseCumulativeNormal ic;
		BoostThreefryUniformRng rng(seed);
		rng.restart(0);		// fixed it
		rng.discard(skip);
		
		for (uint64_t i = 0; i < length; i++)
		{
			buffer[i] = ic(rng.nextReal());
		}
	}

	// Returns normal rsg, caller responsible for array size
	/*
	void normalthreefry2(uint64_t seed, uint64_t skip, double buffer[], uint64_t length)
	{
		boost::random::counter_based_engine<uint64_t, boost::random::threefry<2, uint64_t>, 32> rng(seed);
		rng.restart({ 0, 0 });		// fixed it
		rng.discard(skip);

		boost::normal_distribution <> norm(0, 1);
		boost::variate_generator < boost::random::counter_based_engine<uint64_t, boost::random::threefry<2, uint64_t>, 32> &, boost::normal_distribution <>>
			unnorm(rng, norm);

		for (uint64_t i = 0; i < length; i++)
		{
			buffer[i] = unnorm();
		}
	}
	*/

	// Returns normal rsg, caller responsible for array size
	void gammathreefry(uint64_t seed, uint64_t skip, double buffer[], uint64_t length, Real shape, Real scale)
	{
		QLExtension::InverseCumulativeGamma ic(shape, scale);		// (1,1)
		BoostThreefryUniformRng rng(seed);
		rng.restart(0);		// fixed it
		rng.discard(skip);

		for (uint64_t i = 0; i < length; i++)
		{
			buffer[i] = ic(rng.nextReal());
		}
	}
}