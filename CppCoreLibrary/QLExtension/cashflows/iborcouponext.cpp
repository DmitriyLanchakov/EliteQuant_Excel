/****************************** Project Header ******************************\
Project:	      QuantLib Extension (QLExtension)
Author:			  Letian_zj@gmail.com
URL:			  https://github.com/letianzj/QLExtension
Copyright 2014 Letian_zj
This file is part of  QuantLib Extension (QLExtension) Project.
QuantLib Extension (QLExtension) is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation,
either version 3 of the License, or (at your option) any later version.
QuantTrading is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
See the GNU General Public License for more details.
You should have received a copy of the GNU General Public License along with QLExtension.
If not, see http://www.gnu.org/licenses/.
\***************************************************************************/

#include <cashflows/iborcouponext.hpp>
#include <ql/cashflows/couponpricer.hpp>
#include <ql/cashflows/capflooredcoupon.hpp>
#include <cashflows/cashflowvectorsext.hpp>
#include <ql/indexes/interestrateindex.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/termstructures/yieldtermstructure.hpp>

using boost::shared_ptr;

namespace AMP {

    IborCouponExt::IborCouponExt(const Date& paymentDate,
                           Real nominal,
                           const Date& startDate,
                           const Date& endDate,
                           Natural fixingDays,
                           const shared_ptr<IborIndex>& iborIndex,
                           Real gearing,
                           Spread spread,
                           const Date& refPeriodStart,
                           const Date& refPeriodEnd,
                           const DayCounter& dayCounter,
                           bool isInArrears)
    : FloatingRateCoupon(paymentDate, nominal, startDate, endDate,
                         fixingDays, iborIndex, gearing, spread,
                         refPeriodStart, refPeriodEnd,
                         dayCounter, isInArrears),
      iborIndex_(iborIndex) {

        fixingDate_ = fixingDate();

        const Calendar& fixingCalendar = index_->fixingCalendar();
        Natural indexFixingDays = index_->fixingDays();

        fixingValueDate_ = fixingCalendar.advance(
            fixingDate_, indexFixingDays, Days);

        #ifdef QL_USE_INDEXED_COUPON
        fixingEndDate_ = index_->maturityDate(fixingValueDate_);
        #else
        if (isInArrears_)
            fixingEndDate_ = index_->maturityDate(fixingValueDate_);
        else { // par coupon approximation
            Date nextFixingDate = fixingCalendar.advance(
                accrualEndDate_, -static_cast<Integer>(fixingDays_), Days);
            fixingEndDate_ = fixingCalendar.advance(
                nextFixingDate, indexFixingDays, Days);
        }
        #endif

        const DayCounter& dc = index_->dayCounter();
        spanningTime_ = dc.yearFraction(fixingValueDate_,
                                        fixingEndDate_);
        QL_REQUIRE(spanningTime_>0.0,
                   "\n cannot calculate forward rate between " <<
                   fixingValueDate_ << " and " << fixingEndDate_ <<
                   ":\n non positive time (" << spanningTime_ <<
                   ") using " << dc.name() << " daycounter");
    }

    Rate IborCouponExt::indexFixing() const {

        /* instead of just returning index_->fixing(fixingValueDate_)
           its logic is duplicated here using a specialized iborIndex
           forecastFixing overload which
           1) allows to save date/time recalculations, and
           2) takes into account par coupon needs
        */
        Date today = Settings::instance().evaluationDate();

        if (fixingDate_>today)
            return iborIndex_->forecastFixing(fixingValueDate_,
                                              fixingEndDate_,
                                              spanningTime_);

        if (fixingDate_<today ||
            Settings::instance().enforcesTodaysHistoricFixings()) {
            // do not catch exceptions
            Rate result = index_->pastFixing(fixingDate_);
            QL_REQUIRE(result != Null<Real>(),
                       "Missing " << index_->name() << " fixing for " << fixingDate_);
            return result;
        }

        try {
            Rate result = index_->pastFixing(fixingDate_);
            if (result!=Null<Real>())
                return result;
            else
                ;   // fall through and forecast
        } catch (Error&) {
                ;   // fall through and forecast
        }
        return iborIndex_->forecastFixing(fixingValueDate_,
                                          fixingEndDate_,
                                          spanningTime_);
    }

    void IborCouponExt::accept(AcyclicVisitor& v) {
        Visitor<IborCouponExt>* v1 =
            dynamic_cast<Visitor<IborCouponExt>*>(&v);
        if (v1 != 0)
            v1->visit(*this);
        else
            FloatingRateCoupon::accept(v);
    }



	IborLegExt::IborLegExt(const Schedule& resetschedule, const Schedule& paymentschedule,
                     const shared_ptr<IborIndex>& index)
    : resetschedule_(resetschedule), paymentschedule_(paymentschedule), index_(index),
      paymentAdjustment_(Following),
      inArrears_(false), zeroPayments_(false) {}

	IborLegExt& IborLegExt::withNotionals(Real notional) {
        notionals_ = std::vector<Real>(1,notional);
        return *this;
    }

	IborLegExt& IborLegExt::withNotionals(const std::vector<Real>& notionals) {
        notionals_ = notionals;
        return *this;
    }

	IborLegExt& IborLegExt::withPaymentDayCounter(const DayCounter& dayCounter) {
        paymentDayCounter_ = dayCounter;
        return *this;
    }

	IborLegExt& IborLegExt::withPaymentAdjustment(BusinessDayConvention convention) {
        paymentAdjustment_ = convention;
        return *this;
    }

	IborLegExt& IborLegExt::withFixingDays(Natural fixingDays) {
        fixingDays_ = std::vector<Natural>(1,fixingDays);
        return *this;
    }

	IborLegExt& IborLegExt::withFixingDays(const std::vector<Natural>& fixingDays) {
        fixingDays_ = fixingDays;
        return *this;
    }

	IborLegExt& IborLegExt::withGearings(Real gearing) {
        gearings_ = std::vector<Real>(1,gearing);
        return *this;
    }

	IborLegExt& IborLegExt::withGearings(const std::vector<Real>& gearings) {
        gearings_ = gearings;
        return *this;
    }

	IborLegExt& IborLegExt::withSpreads(Spread spread) {
        spreads_ = std::vector<Spread>(1,spread);
        return *this;
    }

	IborLegExt& IborLegExt::withSpreads(const std::vector<Spread>& spreads) {
        spreads_ = spreads;
        return *this;
    }

	IborLegExt& IborLegExt::withCaps(Rate cap) {
        caps_ = std::vector<Rate>(1,cap);
        return *this;
    }

	IborLegExt& IborLegExt::withCaps(const std::vector<Rate>& caps) {
        caps_ = caps;
        return *this;
    }

	IborLegExt& IborLegExt::withFloors(Rate floor) {
        floors_ = std::vector<Rate>(1,floor);
        return *this;
    }

	IborLegExt& IborLegExt::withFloors(const std::vector<Rate>& floors) {
        floors_ = floors;
        return *this;
    }

	IborLegExt& IborLegExt::inArrears(bool flag) {
        inArrears_ = flag;
        return *this;
    }

	IborLegExt& IborLegExt::withZeroPayments(bool flag) {
        zeroPayments_ = flag;
        return *this;
    }

	// instead of calling FloatingLeg, it construct IborCoupon directly
	IborLegExt::operator Leg() const {

        /*Leg leg = FloatingLegExt<IborIndex, IborCouponExt, CappedFlooredIborCoupon>(
                         schedule_, notionals_, index_, paymentDayCounter_,
                         paymentAdjustment_, fixingDays_, gearings_, spreads_,
                         caps_, floors_, inArrears_, zeroPayments_);*/

		Size n = paymentschedule_.size() - 1;
		QL_REQUIRE(!notionals_.empty(), "no notional given");
		QL_REQUIRE(notionals_.size() <= n,
			"too many nominals (" << notionals_.size() <<
			"), only " << n << " required");
		QL_REQUIRE(gearings_.size() <= n,
			"too many gearings (" << gearings_.size() <<
			"), only " << n << " required");
		QL_REQUIRE(spreads_.size() <= n,
			"too many spreads (" << spreads_.size() <<
			"), only " << n << " required");
		QL_REQUIRE(caps_.size() <= n,
			"too many caps (" << caps_.size() <<
			"), only " << n << " required");
		QL_REQUIRE(floors_.size() <= n,
			"too many floors (" << floors_.size() <<
			"), only " << n << " required");
		QL_REQUIRE(!zeroPayments_ || !inArrears_,
			"in-arrears and zero features are not compatible");

		Leg leg; leg.reserve(n);

		// the following is not always correct
		Calendar calendar = paymentschedule_.calendar();

		Date refStart, start, refEnd, end;
		Date lastPaymentDate = calendar.adjust(paymentschedule_.date(n), paymentAdjustment_);

		// for each payment date
		for (Size i = 0; i < n; ++i) {
			refStart = start = paymentschedule_.date(i);
			refEnd = end = paymentschedule_.date(i + 1);

			std::vector<Date> resetdates;
			std::vector<Date>::const_iterator istart = std::find(resetschedule_.begin(), resetschedule_.end(), start);
			std::vector<Date>::const_iterator iend = std::find(resetschedule_.begin(), resetschedule_.end(), end);
			for (std::vector<Date>::const_iterator it = istart+1; it != iend; i++)
			{
				resetdates.push_back(*it);
			}

			Date paymentDate =
				zeroPayments_ ? lastPaymentDate : calendar.adjust(end, paymentAdjustment_);
			if (i == 0 && !paymentschedule_.isRegular(i + 1)) {
				BusinessDayConvention bdc = paymentschedule_.businessDayConvention();
				refStart = calendar.adjust(end - paymentschedule_.tenor(), bdc);
			}
			if (i == n - 1 && !paymentschedule_.isRegular(i + 1)) {
				BusinessDayConvention bdc = paymentschedule_.businessDayConvention();
				refEnd = calendar.adjust(start + paymentschedule_.tenor(), bdc);
			}

			leg.push_back(boost::shared_ptr<CashFlow>(new
				IborCouponExt(
				paymentDate,
				QuantLib::detail::get(notionals_, i, 1.0),
				start, end,
				QuantLib::detail::get(fixingDays_, i, index_->fixingDays()),
				index_,
				QuantLib::detail::get(gearings_, i, 1.0),
				QuantLib::detail::get(spreads_, i, 0.0),
				refStart, refEnd,
				paymentDayCounter_, inArrears_)));
		}

        if (caps_.empty() && floors_.empty() && !inArrears_) {
            shared_ptr<IborCouponPricer> pricer(new BlackIborCouponPricer);
            setCouponPricer(leg, pricer);
        }

        return leg;
    }

}
