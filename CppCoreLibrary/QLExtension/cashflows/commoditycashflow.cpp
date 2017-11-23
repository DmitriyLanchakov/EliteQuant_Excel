#include <cashflows/commoditycashflow.hpp>
#include <ql/patterns/visitor.hpp>
#include <iomanip>

namespace AMP {

    void CommodityCashFlow::accept(AcyclicVisitor& v) {
        Visitor<CommodityCashFlow>* v1 =
            dynamic_cast<Visitor<CommodityCashFlow>*>(&v);
        if (v1 != 0)
            v1->visit(*this);
        else
            CashFlow::accept(v);
    }

    std::ostream& operator<<(std::ostream& out,
                             const CommodityCashFlows& cashFlows) {
        if (cashFlows.size() == 0)
            return out << "no cashflows" << std::endl;
        out << "cashflows" << std::endl;
        std::string currencyCode; //= cashFlows[0]->discountedAmount().currency().code();
        Real totalDiscounted = 0;
        Real totalUndiscounted = 0;
        for (CommodityCashFlows::const_iterator i = cashFlows.begin();
             i != cashFlows.end(); ++i) {
            //const boost::shared_ptr<CommodityCashFlow> cashFlow = *i;
            const boost::shared_ptr<CommodityCashFlow> cashFlow = i->second;
            totalDiscounted += cashFlow->discountedAmount();
            totalUndiscounted += cashFlow->undiscountedAmount();
            //out << io::iso_date(cashFlow->date()) << " " <<
            out << io::iso_date(i->first) << " "
                << std::setw(16) << std::right << std::fixed
                << std::setprecision(2) << cashFlow->discountedAmount()
                << " " << currencyCode <<
                std::setw(16) << std::right << std::fixed
                << std::setprecision(2)
                << cashFlow->undiscountedAmount() << " " <<
                currencyCode << std::endl;
        }
        out << "total      "
            << std::setw(16) << std::right << std::fixed
            << std::setprecision(2) << totalDiscounted << " " << currencyCode
            << std::setw(16) << std::right << std::fixed
            << std::setprecision(2) << totalUndiscounted << " "
            << currencyCode << std::endl;
        return out;
    }

}