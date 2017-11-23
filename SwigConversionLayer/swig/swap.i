/*
 Copyright (C) 2000, 2001, 2002, 2003 RiskMap srl
 Copyright (C) 2007 StatPro Italia srl
 Copyright (C) 2011 Lluis Pujol Bajador

 This file is part of QuantLib, a free-software/open-source library
 for financial quantitative analysts and developers - http://quantlib.org/

 QuantLib is free software: you can redistribute it and/or modify it
 under the terms of the QuantLib license.  You should have received a
 copy of the license along with this program; if not, please email
 <quantlib-dev@lists.sf.net>. The license is also available online at
 <http://quantlib.org/license.shtml>.

 This program is distributed in the hope that it will be useful, but WITHOUT
 ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
 FOR A PARTICULAR PURPOSE.  See the license for more details.
*/

#ifndef quantlib_swap_i
#define quantlib_swap_i

%include instruments.i
%include termstructures.i
%include cashflows.i
%include timebasket.i

%{
using QuantLib::Swap;
using QuantLib::VanillaSwap;
using QuantLib::OvernightIndexedSwap;
using QLExtension::IBOROISBasisSwap;
using QLExtension::GenericSwap;
using QuantLib::DiscountingSwapEngine;

typedef boost::shared_ptr<Instrument> SwapPtr;
typedef boost::shared_ptr<Instrument> VanillaSwapPtr;
typedef boost::shared_ptr<Instrument> OvernightIndexedSwapPtr;
typedef boost::shared_ptr<Instrument> IBOROISBasisSwapPtr;
typedef boost::shared_ptr<Instrument> GenericSwapPtr;
typedef boost::shared_ptr<PricingEngine> DiscountingSwapEnginePtr;
%}

%rename(Swap) SwapPtr;
class SwapPtr : public boost::shared_ptr<Instrument> {
  public:
    %extend {
        SwapPtr(const std::vector<boost::shared_ptr<CashFlow> >& firstLeg,
                const std::vector<boost::shared_ptr<CashFlow> >& secondLeg) {
            return new SwapPtr(new Swap(firstLeg, secondLeg));
        }
        Date startDate() {
            return boost::dynamic_pointer_cast<Swap>(*self)->startDate();
        }
        Date maturityDate() {
            return boost::dynamic_pointer_cast<Swap>(*self)->maturityDate();
        }
		const Leg & leg(Size i){
            return boost::dynamic_pointer_cast<Swap>(*self)->leg(i);
        }
    }
};


#if defined(SWIGJAVA) || defined(SWIGCSHARP)
%rename(_VanillaSwap) VanillaSwap;
#else
%ignore VanillaSwap;
#endif
class VanillaSwap {
  public:
    enum Type { Receiver = -1, Payer = 1 };
#if defined(SWIGJAVA) || defined(SWIGCSHARP)
  private:
    VanillaSwap();
#endif
};

%rename(VanillaSwap) VanillaSwapPtr;
class VanillaSwapPtr : public SwapPtr {
    #if defined(SWIGMZSCHEME) || defined(SWIGGUILE)
    %rename("fair-rate")        fairRate;
    %rename("fair-spread")      fairSpread;
    %rename("fixed-leg-BPS")    fixedLegBPS;
    %rename("floating-leg-BPS") floatingLegBPS;
	%rename ("fixed-leg-NPV") fixedLegNPV;
	%rename ("floating-leg-NPV") floatingLegNPV;
	%rename ("floating-leg") floatingLeg;
	%rename ("fixed-leg") fixedLeg;
	%rename ("fixed-schedule") fixedSchedule;
	%rename ("floating-schedule") floatingSchedule;
	%rename ("fixed-rate") fixedRate;
	%rename ("fixed-day-count") fixedDayCount;
	%rename ("floating-day-count") floatingDayCount;
	
    #endif
  public:
    %extend {
        static const VanillaSwap::Type Receiver = VanillaSwap::Receiver;
        static const VanillaSwap::Type Payer = VanillaSwap::Payer;
        VanillaSwapPtr(VanillaSwap::Type type, Real nominal,
                       const Schedule& fixedSchedule, Rate fixedRate,
                       const DayCounter& fixedDayCount,
                       const Schedule& floatSchedule,
                       const IborIndexPtr& index,
                       Spread spread,
                       const DayCounter& floatingDayCount) {
            boost::shared_ptr<IborIndex> libor =
                boost::dynamic_pointer_cast<IborIndex>(index);
            return new VanillaSwapPtr(
                    new VanillaSwap(type, nominal,fixedSchedule,fixedRate,
                                    fixedDayCount,floatSchedule,libor,
                                    spread, floatingDayCount));
        }
        Rate fairRate() {
            return boost::dynamic_pointer_cast<VanillaSwap>(*self)->fairRate();
        }
        Spread fairSpread() {
            return boost::dynamic_pointer_cast<VanillaSwap>(*self)
                 ->fairSpread();
        }
        Real fixedLegBPS() {
            return boost::dynamic_pointer_cast<VanillaSwap>(*self)
                 ->fixedLegBPS();
        }
        Real floatingLegBPS() {
            return boost::dynamic_pointer_cast<VanillaSwap>(*self)
                 ->floatingLegBPS();
        }
		Real fixedLegNPV() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->fixedLegNPV();
        }
        Real floatingLegNPV() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->floatingLegNPV();
        }
        // Inspectors 
        const Leg& fixedLeg() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->fixedLeg();
        }
        const Leg& floatingLeg() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->floatingLeg();
        }
        Real nominal() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->nominal();
        }
        const Schedule& fixedSchedule() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->fixedSchedule();
        }
        const Schedule& floatingSchedule() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->floatingSchedule();
        }
        Rate fixedRate() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->fixedRate();
        }
        Spread spread() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->spread();
        }
        const DayCounter& floatingDayCount() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->floatingDayCount();
        }
        const DayCounter& fixedDayCount() {
	        return boost::dynamic_pointer_cast<VanillaSwap> (*self)
		        ->fixedDayCount();
        }
    }
};

#if defined(SWIGJAVA) || defined(SWIGCSHARP)
%rename(_OvernightIndexedSwap) OvernightIndexedSwap;
#else
%ignore OvernightIndexedSwap;
#endif
class OvernightIndexedSwap {
  public:
    enum Type { Receiver = -1, Payer = 1 };
#if defined(SWIGJAVA) || defined(SWIGCSHARP)
  private:
    OvernightIndexedSwap();
#endif
};

%rename(OvernightIndexedSwap) OvernightIndexedSwapPtr;
class OvernightIndexedSwapPtr : public SwapPtr {
    #if defined(SWIGMZSCHEME) || defined(SWIGGUILE)
    %rename("fair-rate")        fairRate;
    %rename("fair-spread")      fairSpread;
    %rename("fixed-leg-BPS")    fixedLegBPS;
    %rename("floating-leg-BPS") floatingLegBPS;
    #endif
  public:
    %extend {
        static const OvernightIndexedSwap::Type Receiver = OvernightIndexedSwap::Receiver;
        static const OvernightIndexedSwap::Type Payer = OvernightIndexedSwap::Payer;
        OvernightIndexedSwapPtr(OvernightIndexedSwap::Type type, Real nominal,
                       const Schedule& schedule, Rate fixedRate,
                       const DayCounter& fixedDC,
                       const OvernightIndexPtr& overnightIndex,
                       Spread spread = 0.0) {
            boost::shared_ptr<OvernightIndex> ois =
                boost::dynamic_pointer_cast<OvernightIndex>(overnightIndex);
            return new OvernightIndexedSwapPtr(
                    new OvernightIndexedSwap(type, nominal,schedule,fixedRate,
                                    fixedDC, ois, spread));
        }
        Rate fairRate() {
            return boost::dynamic_pointer_cast<OvernightIndexedSwap>(*self)->fairRate();
        }
        Spread fairSpread() {
            return boost::dynamic_pointer_cast<OvernightIndexedSwap>(*self)
                 ->fairSpread();
        }
        Real fixedLegBPS() {
            return boost::dynamic_pointer_cast<OvernightIndexedSwap>(*self)
                 ->fixedLegBPS();
        }
        Real overnightLegBPS() {
            return boost::dynamic_pointer_cast<OvernightIndexedSwap>(*self)
                 ->overnightLegBPS();
        }
    }
};

/********************************** Basis swaps ******************************/
#if defined(SWIGJAVA) || defined(SWIGCSHARP)
%rename(_IBOROISBasisSwap) IBOROISBasisSwap;
#else
%ignore IBOROISBasisSwap;
#endif
class IBOROISBasisSwap {
  public:
    enum Type { Receiver = -1, Payer = 1 };
#if defined(SWIGJAVA) || defined(SWIGCSHARP)
  private:
    IBOROISBasisSwap();
#endif
};

%rename(IBOROISBasisSwap) IBOROISBasisSwapPtr;
class IBOROISBasisSwapPtr : public SwapPtr {
    #if defined(SWIGMZSCHEME) || defined(SWIGGUILE)
    %rename("fair-rate")        fairRate;
    %rename("fair-spread")      fairSpread;
    %rename("fixed-leg-BPS")    fixedLegBPS;
    %rename("floating-leg-BPS") floatingLegBPS;
    #endif
  public:
    %extend {
        static const IBOROISBasisSwap::Type Receiver = IBOROISBasisSwap::Receiver;
        static const IBOROISBasisSwap::Type Payer = IBOROISBasisSwap::Payer;
        IBOROISBasisSwapPtr(IBOROISBasisSwap::Type type, Real nominal,
                       const Schedule& floatingSchedule, 
					   const IborIndexPtr& iborIndex,
					   const DayCounter& floatingDC,
					   const Schedule& fixedSchedule,
					   const OvernightIndexPtr& overnightIndex,
					   Spread spread,
					   const DayCounter& fixedDC,
					   BusinessDayConvention paymentConvention = ModifiedFollowing,
                       bool arithmeticAveragedCoupon = true) {
			boost::shared_ptr<IborIndex> ibor =
                boost::dynamic_pointer_cast<IborIndex>(iborIndex);
            boost::shared_ptr<OvernightIndex> ois =
                boost::dynamic_pointer_cast<OvernightIndex>(overnightIndex);
            return new IBOROISBasisSwapPtr(
                    new IBOROISBasisSwap(type, nominal, floatingSchedule, ibor, floatingDC,
							fixedSchedule, ois, spread, fixedDC,
							paymentConvention, arithmeticAveragedCoupon));
        }
        Spread fairSpread() {
            return boost::dynamic_pointer_cast<IBOROISBasisSwap>(*self)
                 ->fairSpread();
        }
        Real floatingLegBPS() {
            return boost::dynamic_pointer_cast<IBOROISBasisSwap>(*self)
                 ->floatingLegBPS();
        }
		Real floatingLegNPV() {
            return boost::dynamic_pointer_cast<IBOROISBasisSwap>(*self)
                 ->floatingLegNPV();
        }
        Real overnightLegBPS() {
            return boost::dynamic_pointer_cast<IBOROISBasisSwap>(*self)
                 ->overnightLegBPS();
        }
		Real overnightLegNPV() {
            return boost::dynamic_pointer_cast<IBOROISBasisSwap>(*self)
                 ->overnightLegNPV();
        }
    }
};

#if defined(SWIGJAVA) || defined(SWIGCSHARP)
%rename(_GenericSwap) GenericSwap;
#else
%ignore GenericSwap;
#endif
class GenericSwap {
  public:
    enum Type { Receiver = -1, Payer = 1 };
#if defined(SWIGJAVA) || defined(SWIGCSHARP)
  private:
    GenericSwap();
#endif
};

%rename(GenericSwap) GenericSwapPtr;
class GenericSwapPtr : public SwapPtr {
    #if defined(SWIGMZSCHEME) || defined(SWIGGUILE)
    %rename("fair-rate")        fairRate;
    %rename("fair-spread")      fairSpread;
    %rename("fixed-leg-BPS")    fixedLegBPS;
    %rename("floating-leg-BPS") floatingLegBPS;
    #endif
  public:
    %extend {
        static const GenericSwap::Type Receiver = GenericSwap::Receiver;
        static const GenericSwap::Type Payer = GenericSwap::Payer;
        GenericSwapPtr(GenericSwap::Type type,
					std::vector<Real> firstLegNominals,
					const Schedule& firstLegResetSchedule,
					const Schedule& firstLegPaymentSchedule,
					const Rate firstLegRate,					// fixed rate
					const DayCounter& firstLegDayCount,
					std::vector<Real> secondLegNominals,
					const Schedule& secondLegResetSchedule,
					const Schedule& secondLegPaymentSchedule,
					const IborIndexPtr& secondLegIndex,
					const DayCounter& secondLegDayCount,
					Spread secondLegSpread = 0.0,
					bool arithmeticAveragedCoupon = false) {
			boost::shared_ptr<IborIndex> ibor2 =
                boost::dynamic_pointer_cast<IborIndex>(secondLegIndex);
            return new GenericSwapPtr(
                    new GenericSwap(type, firstLegNominals, firstLegResetSchedule, firstLegPaymentSchedule, firstLegRate, firstLegDayCount,
							secondLegNominals, secondLegResetSchedule, secondLegPaymentSchedule, ibor2, secondLegDayCount, secondLegSpread,
							arithmeticAveragedCoupon));
        }
		GenericSwapPtr(GenericSwap::Type type,
					std::vector<Real> firstLegNominals,
					const Schedule& firstLegResetSchedule,
					const Schedule& firstLegPaymentSchedule,
					const IborIndexPtr& firstLegIndex,					// floating index
					const DayCounter& firstLegDayCount,
					std::vector<Real> secondLegNominals,
					const Schedule& secondLegResetSchedule,
					const Schedule& secondLegPaymentSchedule,
					const IborIndexPtr& secondLegIndex,
					const DayCounter& secondLegDayCount,
					Spread firstLegSpread = 0.0,
					Spread secondLegSpread = 0.0,
					bool arithmeticAveragedCoupon = false) {
			boost::shared_ptr<IborIndex> ibor1 =
                boost::dynamic_pointer_cast<IborIndex>(firstLegIndex);
			boost::shared_ptr<IborIndex> ibor2 =
                boost::dynamic_pointer_cast<IborIndex>(secondLegIndex);
            return new GenericSwapPtr(
                    new GenericSwap(type, firstLegNominals, firstLegResetSchedule, firstLegPaymentSchedule, ibor1, firstLegDayCount,
							secondLegNominals, secondLegResetSchedule, secondLegPaymentSchedule, ibor2, secondLegDayCount, 
							firstLegSpread, secondLegSpread,
							arithmeticAveragedCoupon));
        }
		
		GenericSwap::Type type() {
			return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->type();
		}
		
		Real fairRate() {
            return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->fairRate();
        }
        Spread fairSpread() {
            return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->fairSpread();
        }
        Real firstLegBPS() {
            return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->firstLegBPS();
        }
		Real firstLegNPV() {
            return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->firstLegNPV();
        }
        Real secondLegBPS() {
            return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->secondLegBPS();
        }
		Real secondLegNPV() {
            return boost::dynamic_pointer_cast<GenericSwap>(*self)
                 ->secondLegNPV();
        }
		/*const std::vector<boost::shared_ptr<CashFlow> >& leg(Size j) {
			return boost::dynamic_pointer_cast<GenericSwap>(*self)->leg(j);
		}*/
		const std::vector<std::string>& firstLegInfo() {
			return boost::dynamic_pointer_cast<GenericSwap>(*self)->firstLegInfo();
		}
		const std::vector<std::string>& secondLegInfo() {
			return boost::dynamic_pointer_cast<GenericSwap>(*self)->secondLegInfo();
		}
    }
};

/********************************** End Basis Swaps **************************/
%rename(DiscountingSwapEngine) DiscountingSwapEnginePtr;
class DiscountingSwapEnginePtr : public boost::shared_ptr<PricingEngine> {
  public:
    %extend {
        DiscountingSwapEnginePtr(
                            const Handle<YieldTermStructure>& discountCurve,
                            const Date& settlementDate = Date(),
                            const Date& npvDate = Date()) {
            return new DiscountingSwapEnginePtr(
                                    new DiscountingSwapEngine(discountCurve,
                                                              boost::none,
                                                              settlementDate,
                                                              npvDate));
        }
        DiscountingSwapEnginePtr(
                            const Handle<YieldTermStructure>& discountCurve,
                            bool includeSettlementDateFlows,
                            const Date& settlementDate = Date(),
                            const Date& npvDate = Date()) {
            return new DiscountingSwapEnginePtr(
                         new DiscountingSwapEngine(discountCurve,
                                                   includeSettlementDateFlows,
                                                   settlementDate,
                                                   npvDate));
        }
    }
};


%{
using QuantLib::AssetSwap;
typedef boost::shared_ptr<Instrument> AssetSwapPtr;
%}

%rename(AssetSwap) AssetSwapPtr;
class AssetSwapPtr : public SwapPtr {
    #if !defined(SWIGJAVA) && !defined(SWIGCSHARP)
    %feature("kwargs") AssetSwapPtr;
    #endif
  public:
    %extend {
        AssetSwapPtr(bool payFixedRate,
                     const BondPtr& bond,
                     Real bondCleanPrice,
                     const InterestRateIndexPtr& index,
                     Spread spread,
                     const Schedule& floatSchedule = Schedule(),
                     const DayCounter& floatingDayCount = DayCounter(),
                     bool parAssetSwap = true) {
            const boost::shared_ptr<Bond> b =
                boost::dynamic_pointer_cast<Bond>(bond);
            const boost::shared_ptr<IborIndex> i =
                boost::dynamic_pointer_cast<IborIndex>(index);
            return new AssetSwapPtr(
                new AssetSwap(payFixedRate,b,bondCleanPrice,i,spread,
                              floatSchedule,floatingDayCount,parAssetSwap));
        }
        Real fairCleanPrice() {
            return boost::dynamic_pointer_cast<AssetSwap>(*self)
                ->fairCleanPrice();
        }
        Spread fairSpread() {
            return boost::dynamic_pointer_cast<AssetSwap>(*self)
                ->fairSpread();
        }
    }
};

/************************ commodity **************/
// from class index, for boost::shared_ptr<base> with constructor
%{
using QLExtension::PricingPeriodExt;
typedef boost::shared_ptr<QLExtension::PricingPeriodExt> PricingPeriodExtPtr;
%}

%rename(PricingPeriodExt) PricingPeriodExtPtr;
class PricingPeriodExtPtr {
  public:
    %extend {
        PricingPeriodExtPtr(const Date& startDate, const Date& endDate,
                      const Date& paymentDate, const Real quantity, 
					  const Real payCoeff = 1, const Real recCoeff = 1, 
					  const Real paySpread = 0, const Real recSpread = 0) {
            return new PricingPeriodExtPtr(new PricingPeriodExt(startDate, endDate, paymentDate, quantity,
					payCoeff, recCoeff, paySpread, recSpread));
		}
		Date paymentDate() {
			return (*self)->paymentDate();
		}
		Real quantity() {
			return (*self)->quantity();
		}
		Real payCoeff() { 
			return (*self)->payCoeff(); 
		}
		Real paySpread() { 
			return (*self)->paySpread(); 
		}
		Real recCoeff() { 
			return (*self)->recCoeff(); 
		}
		Real recSpread() { 
			return (*self)->recSpread(); 
		}
		Real getuPayDelta() { 
			return (*self)->getuPayDelta(); 
		}
		Real getdPayDelta() { 
			return (*self)->getdPayDelta(); 
		}
		Real getuRecDelta() { 
			return (*self)->getuRecDelta(); 
		}
		Real getdRecDelta() { 
			return (*self)->getdRecDelta(); 
		}
		Real getFinalized() { 
			return (*self)->getFinalized(); 
		}
    }
};

#if defined(SWIGCSHARP)
SWIG_STD_VECTOR_ENHANCED( PricingPeriodExtPtr )
#endif
%template(PricingPeriodExts) std::vector<PricingPeriodExtPtr >;
typedef std::vector<PricingPeriodExtPtr > PricingPeriodExts;

%{
using QLExtension::EnergyCommodityExt;
using QLExtension::EnergyFutureExt;
using QLExtension::EnergySwapExt;
using QLExtension::EnergyVanillaSwapExt;
using QLExtension::EnergyBasisSwapExt;

typedef boost::shared_ptr<Instrument> EnergyCommodityExtPtr;
typedef boost::shared_ptr<Instrument> EnergyFutureExtPtr;
typedef boost::shared_ptr<Instrument> EnergySwapExtPtr;
typedef boost::shared_ptr<Instrument> EnergyVanillaSwapExtPtr;
typedef boost::shared_ptr<Instrument> EnergyBasisSwapExtPtr;
%}

%rename(EnergyCommodityExt) EnergyCommodityExtPtr;
class EnergyCommodityExtPtr : public boost::shared_ptr<Instrument> {
public:
    %extend {
	std::string commodityName() {
            return boost::dynamic_pointer_cast<EnergyCommodityExt>(*self)->commodityName();
        }
	}
};

#if defined(SWIGJAVA) || defined(SWIGCSHARP)
%rename(_EnergyFutureExt) EnergyFutureExt;
#else
%ignore EnergyFutureExt;
#endif
class EnergyFutureExt {
#if defined(SWIGJAVA) || defined(SWIGCSHARP)
  private:
    EnergyFutureExt();
#endif
};

%rename(EnergyFutureExt) EnergyFutureExtPtr;
class EnergyFutureExtPtr : public EnergyCommodityExtPtr {
  public:
    %extend {
        EnergyFutureExtPtr(Integer buySell, const PricingPeriodExtPtr& PricingPeriodExt,
                     const Real tradePrice,
                     const boost::shared_ptr<CommodityIndexExt>& index,
                     const std::string& commodityName,
					 const Handle<YieldTermStructure>& discountTermStructure) {
            return new EnergyFutureExtPtr(
                    new EnergyFutureExt(buySell, PricingPeriodExt, tradePrice, index, commodityName, discountTermStructure));
        }
        PricingPeriodExtPtr pricingPeriod() {
            return boost::dynamic_pointer_cast<EnergyFutureExt>(*self)->pricingPeriod();
        }
        Real tradePrice() {
            return boost::dynamic_pointer_cast<EnergyFutureExt>(*self)->tradePrice();
        }
		Real quantity() {
            return boost::dynamic_pointer_cast<EnergyFutureExt>(*self)->quantity();
        }
        boost::shared_ptr<CommodityIndexExt> index() {
            return boost::dynamic_pointer_cast<EnergyFutureExt>(*self)->index();
        }
    }
};

%rename(EnergySwapExt) EnergySwapExtPtr;
class EnergySwapExtPtr : public EnergyCommodityExtPtr {
  public:
    %extend {
        EnergySwapExtPtr(const std::vector<PricingPeriodExtPtr >& PricingPeriodExts,
                   const std::string& commodityName,
				   const Frequency deliverySchedule,
				   const Calendar& calendar) {
            return new EnergySwapExtPtr(
                    new EnergySwapExt(PricingPeriodExts, commodityName, deliverySchedule, calendar));
        }
        /*bool isExpired() {
            return boost::dynamic_pointer_cast<EnergySwapExt>(*self)->isExpired();
        }*/
        std::vector<PricingPeriodExtPtr > pricingPeriods() {
            return boost::dynamic_pointer_cast<EnergySwapExt>(*self)->pricingPeriods();
        }
        const std::string commodityName() {
            return boost::dynamic_pointer_cast<EnergySwapExt>(*self)->commodityName();
        }
		Real quantity() {
            return boost::dynamic_pointer_cast<EnergySwapExt>(*self)->quantity();
        }
    }
};

%rename(EnergyVanillaSwapExt) EnergyVanillaSwapExtPtr;
class EnergyVanillaSwapExtPtr : public EnergySwapExtPtr {
  public:
    %extend {
        EnergyVanillaSwapExtPtr(bool payer,
                    const Real fixedPrice,
					const boost::shared_ptr<CommodityIndexExt>& index,
                    const std::vector<PricingPeriodExtPtr >& PricingPeriodExts,
                    const std::string& commodityName,
                    const Handle<YieldTermStructure>& payLegTermStructure,
                    const Handle<YieldTermStructure>& receiveLegTermStructure,
					const Frequency deliverySchedule = Frequency::Daily,
					const Calendar& calendar = QuantLib::NullCalendar()) {
            return new EnergyVanillaSwapExtPtr(
                    new EnergyVanillaSwapExt(payer, fixedPrice, index, PricingPeriodExts,
						commodityName, payLegTermStructure, receiveLegTermStructure,
						deliverySchedule, calendar));
        }
        boost::shared_ptr<CommodityIndexExt> index() {
            return boost::dynamic_pointer_cast<EnergyVanillaSwapExt>(*self)->index();
        }
        Integer payReceive() {
            return boost::dynamic_pointer_cast<EnergyVanillaSwapExt>(*self)->payReceive();
        }
		Real fixedPrice() {
            return boost::dynamic_pointer_cast<EnergyVanillaSwapExt>(*self)->fixedPrice();
        }
    }
};

%rename(EnergyBasisSwapExt) EnergyBasisSwapExtPtr;
class EnergyBasisSwapExtPtr : public EnergySwapExtPtr {
  public:
    %extend {
        EnergyBasisSwapExtPtr(const boost::shared_ptr<CommodityIndexExt>& payIndex,
					const boost::shared_ptr<CommodityIndexExt>& receiveIndex,
                    const std::vector<PricingPeriodExtPtr >& PricingPeriodExts,			// shared by pay/rec legs
                    const std::string& commodityName,
                    const Handle<YieldTermStructure>& payLegTermStructure,
                    const Handle<YieldTermStructure>& receiveLegTermStructure,
					const Frequency deliverySchedule = Frequency::Daily,
					const Calendar& calendar = QuantLib::NullCalendar()) {
            return new EnergyBasisSwapExtPtr(
                    new EnergyBasisSwapExt(payIndex, receiveIndex, PricingPeriodExts,
						commodityName, payLegTermStructure,
						receiveLegTermStructure, deliverySchedule, calendar));
        }
        boost::shared_ptr<CommodityIndexExt> payIndex() {
            return boost::dynamic_pointer_cast<EnergyBasisSwapExt>(*self)->payIndex();
        }
		boost::shared_ptr<CommodityIndexExt> receiveIndex() {
            return boost::dynamic_pointer_cast<EnergyBasisSwapExt>(*self)->receiveIndex();
        }
    }
};

#endif
