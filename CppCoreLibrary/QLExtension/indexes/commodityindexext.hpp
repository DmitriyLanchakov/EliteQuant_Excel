#ifndef qlex_commodity_index_ext_hpp
#define qlex_commodity_index_ext_hpp

#include <termstructures/commodity/commoditycurveext.hpp>
#include <ql/indexes/indexmanager.hpp>

namespace QLExtension {

    class TermStructure;

    //! base class for commodity indexes
    class CommodityIndexExt : public Observable,
                           public Observer {
      public:
		CommodityIndexExt(
                const std::string& name,
				Real lotQuantity = 1.0,
				const Calendar& calendar = QuantLib::NullCalendar());
        CommodityIndexExt(
                const std::string& name,
				const boost::shared_ptr<CommodityCurveExt>& forwardCurve,
				Real lotQuantity = 1.0,
				const Calendar& calendar = QuantLib::NullCalendar());
        //! \name Index interface
        //@{
        std::string name() const;
        //@}
        //! \name Observer interface
        //@{
        void update();
        //@}
        //! \name Inspectors
        //@{
        const Calendar& calendar() const;
		const boost::shared_ptr<CommodityCurveExt>& forwardCurve() const;
		void setForwardCurve(const boost::shared_ptr<CommodityCurveExt>& forwardCurve);
        Real lotQuantity() const;

        Real price(const Date& date);			// (historical) spot price
        Real forwardPrice(const Date& date) const;
        Date lastQuoteDate() const;
        //@}
        void addQuote(const Date& quoteDate, Real quote);

		// add historical date's quotes
        void addQuotes(const std::map<Date, Real>& quotes) {
            std::string tag = name();
            quotes_ = IndexManager::instance().getHistory(tag);
            for (std::map<Date, Real>::const_iterator ii = quotes.begin();
                 ii != quotes.end (); ii++) {
                quotes_[ii->first] = ii->second;
            }
            IndexManager::instance().setHistory(tag, quotes_);
        }

        void clearQuotes();
        //! returns TRUE if the quote date is valid
        bool isValidQuoteDate(const Date& quoteDate) const;
        bool empty() const;
        bool forwardCurveEmpty() const;
        const TimeSeries<Real>& quotes() const;

        friend std::ostream& operator<<(std::ostream&, const CommodityIndexExt&);
      protected:
        std::string name_;
        Calendar calendar_;
        Real lotQuantity_;
        TimeSeries<Real> quotes_;
		boost::shared_ptr<CommodityCurveExt> forwardCurve_;
    };


    // inline definitions

    inline bool operator==(const CommodityIndexExt& i1, const CommodityIndexExt& i2) {
        return i1.name() == i2.name();
    }

    inline void CommodityIndexExt::update() {
        notifyObservers();
    }

    inline std::string CommodityIndexExt::name() const {
        return name_;
    }

    inline const Calendar& CommodityIndexExt::calendar() const {
        return calendar_;
    }

	inline Real CommodityIndexExt::lotQuantity() const {
        return lotQuantity_;
    }

	inline const boost::shared_ptr<CommodityCurveExt>&
		CommodityIndexExt::forwardCurve() const {
        return forwardCurve_;
    }

	inline const TimeSeries<Real>& CommodityIndexExt::quotes() const {
        return quotes_;
    }

	// historical quotes
	// bug: quotes_find automatically adds new values to quotes_
	// this skips IndexManager
	// so to call price, it has to make sure dates are in the fixing dates
	inline Real CommodityIndexExt::price(const Date& date) {
        std::map<Date, Real>::const_iterator hq = quotes_.find(date);
        if (hq->second == Null<Real>()) {
            hq++;
            if (hq == quotes_.end())
                //if (hq->second == Null<Real>())
                return Null<Real>();
        }
        return hq->second;
    }

	inline Real CommodityIndexExt::forwardPrice(const Date& date) const {
        try {
            Real forwardPrice = forwardCurve_->price(date);
            return forwardPrice;
        } catch (const std::exception& e) {
            QL_FAIL("error fetching forward price for index " << name_
                    << ": " << e.what());
        }
    }

	inline Date CommodityIndexExt::lastQuoteDate() const {
        if (quotes_.empty())
            return Date::minDate();
        return quotes_.lastDate();
    }

	inline bool CommodityIndexExt::empty() const {
        return quotes_.empty();
    }

	inline bool CommodityIndexExt::forwardCurveEmpty() const {
        if (forwardCurve_ != 0)
            return forwardCurve_->empty();
        return false;
    }

	inline void CommodityIndexExt::addQuote(const Date& quoteDate, Real quote) {
        //QL_REQUIRE(isValidQuoteDate(quoteDate),
        //           "Quote date " << quoteDate.weekday() << ", " <<
        //           quoteDate << " is not valid");
        std::string tag = name();
        quotes_ = IndexManager::instance().getHistory(tag);
        quotes_[quoteDate] = quote;
        IndexManager::instance().setHistory(tag, quotes_);
    }

	inline void CommodityIndexExt::clearQuotes() {
        IndexManager::instance().clearHistory(name());
    }

	inline bool CommodityIndexExt::isValidQuoteDate(const Date& quoteDate) const {
        return calendar().isBusinessDay(quoteDate);
    }

}

#endif
