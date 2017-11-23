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

// 1. ql/indexes/iborindex.hpp: forecastFixing needs to be public (can't add friend)
#ifndef amp_ibor_coupon_ext_hpp
#define amp_ibor_coupon_ext__hpp

#include <ql/cashflows/floatingratecoupon.hpp>
#include <ql/indexes/iborindex.hpp>
#include <ql/time/schedule.hpp>

using namespace QuantLib;

namespace AMP {

    //! %Coupon paying a Libor-type index
    class IborCouponExt : public FloatingRateCoupon {
      public:
        IborCouponExt(const Date& paymentDate,
                   Real nominal,
                   const Date& startDate,
                   const Date& endDate,
                   Natural fixingDays,
                   const boost::shared_ptr<IborIndex>& index,
                   Real gearing = 1.0,
                   Spread spread = 0.0,
                   const Date& refPeriodStart = Date(),
                   const Date& refPeriodEnd = Date(),
                   const DayCounter& dayCounter = DayCounter(),
                   bool isInArrears = false);
        //! \name Inspectors
        //@{
        const boost::shared_ptr<IborIndex>& iborIndex() const {
            return iborIndex_;
        }
        //@}
        //! \name FloatingRateCoupon interface
        //@{
        //! Implemented in order to manage the case of par coupon
        Rate indexFixing() const;
        //@}
        //! \name Visitability
        //@{
        virtual void accept(AcyclicVisitor&);
        //@}
      private:
        boost::shared_ptr<IborIndex> iborIndex_;
        Date fixingDate_, fixingValueDate_, fixingEndDate_;
        Time spanningTime_;
    };


    //! helper class building a sequence of capped/floored ibor-rate coupons
    class IborLegExt {
      public:
        IborLegExt(const Schedule& resetschedule, const Schedule& paymentschedule,
                const boost::shared_ptr<IborIndex>& index);
		IborLegExt& withNotionals(Real notional);
		IborLegExt& withNotionals(const std::vector<Real>& notionals);
		IborLegExt& withPaymentDayCounter(const DayCounter&);
		IborLegExt& withPaymentAdjustment(BusinessDayConvention);
		IborLegExt& withFixingDays(Natural fixingDays);
		IborLegExt& withFixingDays(const std::vector<Natural>& fixingDays);
		IborLegExt& withGearings(Real gearing);
		IborLegExt& withGearings(const std::vector<Real>& gearings);
		IborLegExt& withSpreads(Spread spread);
		IborLegExt& withSpreads(const std::vector<Spread>& spreads);
		IborLegExt& withCaps(Rate cap);
		IborLegExt& withCaps(const std::vector<Rate>& caps);
		IborLegExt& withFloors(Rate floor);
		IborLegExt& withFloors(const std::vector<Rate>& floors);
		IborLegExt& inArrears(bool flag = true);
		IborLegExt& withZeroPayments(bool flag = true);
        operator Leg() const;
      private:
        Schedule resetschedule_;
		Schedule paymentschedule_;
        boost::shared_ptr<IborIndex> index_;
        std::vector<Real> notionals_;
        DayCounter paymentDayCounter_;
        BusinessDayConvention paymentAdjustment_;
        std::vector<Natural> fixingDays_;
        std::vector<Real> gearings_;
        std::vector<Spread> spreads_;
        std::vector<Rate> caps_, floors_;
        bool inArrears_, zeroPayments_;
    };

}

#endif
