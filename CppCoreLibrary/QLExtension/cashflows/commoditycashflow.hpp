#ifndef amp_commodity_cash_flow_hpp
#define amp_commodity_cash_flow_hpp

#include <ql/cashflow.hpp>
#include <ql/money.hpp>
#include <map>

using namespace QuantLib;

namespace AMP {

    class CommodityCashFlow : public CashFlow {
      public:
        CommodityCashFlow(const Date& date,
                          Real discountedAmount,
                          Real undiscountedAmount,
                          Real discountedPaymentAmount,
                          Real undiscountedPaymentAmount,
                          Real discountFactor,
                          Real paymentDiscountFactor,
                          bool finalized)
        : date_(date), discountedAmount_(discountedAmount),
          undiscountedAmount_(undiscountedAmount),
          discountedPaymentAmount_(discountedPaymentAmount),
          undiscountedPaymentAmount_(undiscountedPaymentAmount),
          discountFactor_(discountFactor),
          paymentDiscountFactor_(paymentDiscountFactor),
          finalized_(finalized) {}
        //! \name Event interface
        //@{
        Date date() const { return date_; }
        //@}
        //! \name CashFlow interface
        //@{
        Real amount() const { return discountedAmount_; }
        //@}

		const Real discountedAmount() const { return discountedAmount_; }
		const Real undiscountedAmount() const { return undiscountedAmount_; }
        const Real discountedPaymentAmount() const {
            return discountedPaymentAmount_;
        }
        const Real undiscountedPaymentAmount() const {
            return undiscountedPaymentAmount_;
        }
        Real discountFactor() const { return discountFactor_; }
        Real paymentDiscountFactor() const { return paymentDiscountFactor_; }
        bool finalized() const { return finalized_; }

        //! \name Visitability
        //@{
        virtual void accept(AcyclicVisitor&);
        //@}
      private:
        Date date_;
        Real discountedAmount_, undiscountedAmount_,
              discountedPaymentAmount_, undiscountedPaymentAmount_;
        Real discountFactor_, paymentDiscountFactor_;
        bool finalized_;
    };

    typedef std::map<Date, boost::shared_ptr<CommodityCashFlow> >
                                                           CommodityCashFlows;

    #ifndef __DOXYGEN__
    std::ostream& operator<<(std::ostream& out,
                             const CommodityCashFlows& cashFlows);
    #endif

}

#endif
