#ifndef qlex_overnight_indexed_coupon_hpp
#define qlex_overnight_indexed_coupon_hpp

#include <ql/cashflows/floatingratecoupon.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/time/schedule.hpp>
#include <ql/cashflows/overnightindexedcoupon.hpp>
#include <ql/cashflows/couponpricer.hpp>
#include <ql/quotes/simplequote.hpp>

using namespace QuantLib;

namespace QLExtension {
	/*! pricer for arithmetically averaged overnight indexed coupons
		Reference: Katsumi Takada 2011, Valuation of Arithmetically Average of Fed Funds Rates and Construction of the US Dollar Swap Yield Curve
	*/
	class ArithmeticAveragedOvernightIndexedCouponPricer : public FloatingRateCouponPricer {
		public:
			ArithmeticAveragedOvernightIndexedCouponPricer(Handle<Quote> meanReversion = Handle<Quote>(boost::shared_ptr<Quote>(new SimpleQuote(0.03))),
				Handle<Quote> vol = Handle<Quote>(boost::shared_ptr<Quote>(new SimpleQuote(0.00))))
				 : meanReversion_(meanReversion), vol_(vol) {}
			void initialize(const FloatingRateCoupon& coupon);
			Rate swapletRate() const;
			
				Real swapletPrice() const { QL_FAIL("swapletPrice not available"); }
			Real capletPrice(Rate) const { QL_FAIL("capletPrice not available"); }
			Rate capletRate(Rate) const { QL_FAIL("capletRate not available"); }
			Real floorletPrice(Rate) const { QL_FAIL("floorletPrice not available"); }
			Rate floorletRate(Rate) const { QL_FAIL("floorletRate not available"); }
			
				Real meanReversion() const { meanReversion_->value(); };
			Real volatility() const { vol_->value(); };
		protected:
			Real convAdj1(Time ts, Time te) const;
			Real convAdj2(Time ts, Time te) const;
			const OvernightIndexedCoupon* coupon_;
			Handle<Quote> meanReversion_;
			Handle<Quote> vol_;
	};
}

#endif
