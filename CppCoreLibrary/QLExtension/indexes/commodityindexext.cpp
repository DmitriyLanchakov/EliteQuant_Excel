#include <indexes/commodityindexext.hpp>

namespace QLExtension {

    CommodityIndexExt::CommodityIndexExt(
                const std::string& indexName,
				const boost::shared_ptr<CommodityCurveExt>& forwardCurve,
				Real lotQuantity, 
				const Calendar& calendar)
    : name_(indexName), calendar_(calendar),
      lotQuantity_(lotQuantity), forwardCurve_(forwardCurve)
	{
		if (IndexManager::instance().hasHistory(indexName))
		{
			quotes_ = IndexManager::instance().getHistory(indexName);
			IndexManager::instance().setHistory(indexName, quotes_);
		}
        registerWith(Settings::instance().evaluationDate());
        registerWith(IndexManager::instance().notifier(name()));
    }

	CommodityIndexExt::CommodityIndexExt(
		const std::string& indexName,
		Real lotQuantity,
		const Calendar& calendar)
		: name_(indexName), calendar_(calendar),
		lotQuantity_(lotQuantity)
	{
		if (IndexManager::instance().hasHistory(indexName))
		{
			quotes_ = IndexManager::instance().getHistory(indexName);
			IndexManager::instance().setHistory(indexName, quotes_);
		}
		registerWith(Settings::instance().evaluationDate());
		registerWith(IndexManager::instance().notifier(name()));
	}

	void CommodityIndexExt::setForwardCurve(const boost::shared_ptr<CommodityCurveExt>& forwardCurve)
	{
		if (forwardCurve_ != 0)
			forwardCurve_.reset();		// release old commodity curve

		forwardCurve_ = forwardCurve;
		notifyObservers();
	}

    std::ostream& operator<<(std::ostream& out, const CommodityIndexExt& index) {
        out << "[" << index.name_ << "]";
        if (index.forwardCurve_ != 0)
            out << "; forward (" << (*index.forwardCurve_) << ")";
        return out;
    }

}
