// http://www.boost.org/doc/libs/1_55_0/doc/html/boost_random/reference.html#boost_random.reference.generators
// http://www.johndcook.com/blog/cpp_TR1_random/#gamma
#ifndef qlex_boost_rng_binding_h
#define qlex_boost_rng_binding_h

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
#include <math/distributions/gammadistribution.hpp>
#include <settings.hpp>

using namespace QuantLib;

namespace QLExtension {

	template<class RNG>
	class GenericBoostIntRNG
	{
	private:
		mutable RNG generator;
		static Real ceil_value;
	public:
		typedef Sample<Real> sample_type;
		typedef typename RNG::result_type result_type;

		BOOST_STATIC_ASSERT(std::numeric_limits<result_type>::is_integer);

		GenericBoostIntRNG(result_type seed = 0);

		//! return a random number in the (0.0, 1.0)-interval
		//! For random number generators with a precision larger than
		Real nextReal() const
		{
			return (Real(generator()) + 0.5) / ceil_value;
		}

		void discard(result_type n)
		{
			generator.discard(n);
		}

		// restart/reset base counter
		// seed is fixed to a name/curve. 
		// Use this to restart inside the name. it is intended for use in inner loops. 
		void restart(result_type n)
		{
			generator.restart({ n, 0 });
		}

		/*! returns a sample with weight 1.0 containing a random number
		in the (0.0, 1.0) interval  */
		sample_type next() const
		{
			return sample_type(nextReal(), 1.0);
		}

	private:
		uint64_t getSeed(uint64_t seed);
	};		// GenericBoostIntRNG

	// counter-bits + non-counter-bits (base counter) = domain_type
	// Attempting to set a key (either with the constructor or with seed(...) that has non-zero bits in the highest log_2(DomainBits) of its key 
	//			will throw an invalid_argument exception.

	// Boost threefry 2x64 round 32
	// uint64 --> 64bit base counter
	//  threefry<N, U>::key_type    = boost::array<U, N>
	//					domain_type = boost::array<U, N>
	//					range_type = boost::array<U, N>
	typedef GenericBoostIntRNG< boost::random::counter_based_engine<uint64_t, 
		boost::random::threefry<2, uint64_t>, 32> > BoostThreefryUniformRng;
	typedef GenericPseudoRandom<
		BoostThreefryUniformRng,
		InverseCumulativeNormal> PseudoRandomNormalBoostThreefry;

	// Boost Philox 2x64 round 10
	// uint64 --> 64bit base counter
	// philox<N, U>::key_type    = boost::array<U, N/2>
	//				domain_type = boost::array<U, N>
	//				range_type = boost::array<U, N>
	typedef GenericBoostIntRNG<boost::random::counter_based_engine<uint64_t,
		boost::random::philox<2, uint64_t>> > BoostPhiloxUniformRng;
	typedef GenericPseudoRandom<
		BoostPhiloxUniformRng,
		InverseCumulativeNormal> PseudoRandomNormalBoostPhilox;

	template<class RNG>
	Real GenericBoostIntRNG<RNG>::ceil_value = Real(RNG::max()) + 1.0;

	template<class RNG>
	GenericBoostIntRNG<RNG>::GenericBoostIntRNG(result_type seed) :
		generator(GenericBoostIntRNG<RNG>::getSeed(seed))
	{
	}

	template<class RNG>
	uint64_t GenericBoostIntRNG<RNG>::getSeed(
		uint64_t seed)
	{
		//unsigned long s = (seed != 0 ? seed : SeedGenerator::instance().get());
		//return (result_type)(s % RNG::max());
		return seed;
	}

	void originalthreefry(uint64_t seed, uint64_t skip, uint64_t buffer[], uint64_t length);
	void uniformthreefry(uint64_t seed, uint64_t skip, double buffer[], uint64_t length);
	void normalthreefry(uint64_t seed, uint64_t skip, double buffer[], uint64_t length);
	//QLEX_API void normalthreefry2(uint64_t seed, uint64_t skip, double buffer[], uint64_t length);
	void gammathreefry(uint64_t seed, uint64_t skip, double buffer[], uint64_t length, Real shape, Real scale);
}
#endif /* qlex_boost_rng_binding_h */
