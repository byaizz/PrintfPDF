#ifndef ROLLSCHEDULEDEFINITION_H
#define ROLLSCHEDULEDEFINITION_H

namespace PDFCreater{

	//PDI信息
	struct PDIData{
		char steelGrade[256];//钢种名
		double slabWidthL3;//板坯宽度-L3[mm]
		double slabWidthAct;//板坯宽度-实际[mm]
		double slabThickL3;//板坯厚度-L3[mm]
		double slabThickAct;//板坯厚度-实际[mm]
		double slabLengthL3;//板坯长度-L3[mm]
		double slabLengthAct;//板坯长度-实际[mm]
		double slabWeightL3;//板坯重量-L3
		double slabWeightAct;//板坯重量-实际
		int furnaceNo;//加热炉号
		int furnaceCol;//加热炉号

		double plateThickSet;//成品目标厚度[mm]
		double plateThickAct;//成品实际厚度[mm]
		double plateWidthSet;//成品目标宽度[mm]
		double plateWidthAct;//成品实际宽度[mm]
		double plateLengthSet;//成品目标长度[mm]
		double plateLengthAct;//成品实际长度[mm]
		int finalTemp;//终轧温度
	};

	//轧制设定信息
	struct RollSetup{
		//控扎模式
		enum Enum_CtrlMode{
			NOCTRL,//非控扎
			ONECTRL,//单块控扎
			TWOCTRL,//两块控扎
			THRCTRL,//三块控扎
			CHAINCTRL//链式控扎
		};
		//转钢模式
		enum Enum_TurnMode{
			TM_V,//纵轧
			TM_V_H,//纵-横轧
			TM_V_H_V,//纵-横-纵轧
			TM_H,//横轧
			TM_H_V,//横-纵轧
			TM_H_V_H//横-纵-横轧
		};
		int ctrlMode;//轧制模式，0为RM,1为FM
		int totalPassNum;//总道次数
		int turnMode;//转钢模式
		int turnPassNum;//转钢道次
		double turnThick;//转钢厚度[mm]
		double controlThick;//控扎厚度[mm]
	};

	//轧辊信息
	struct RollData{
		char RollID[256];//轧辊编号
		char RollMaterial[256];//轧辊材质
		char RollStartTime[256];//轧辊直径[mm]
		double RollDiameter;//上机时间
	};

	//轧机轧辊信息
	struct MillData{
		RollData upWorkRoll;//上工作辊信息
		RollData downWorkRoll;//下工作辊信息
		RollData upBackupRoll;//上支撑辊信息
		RollData downBackupRoll;//下支撑辊信
	};

	//道次信息
	struct PassData{
		int millType;//轧机类型(RM/FM)
		int passNo;//道次号
		double passGapSet;//道次辊缝设定值
		double passGapAct;//道次辊缝实际值
		double passThickSet;//道次厚度设定值
		double passThickAct;//道次厚度实际值
		double passWidthSet;//道次宽度设定值
		double passWidthAct;//道次宽度实际值
		double passForceSet;//道次轧制力设定值
		double passForceAct;//道次轧制力实际值
		double passTorqueSet;//道次力矩设定值
		double passTorqueAct;//道次力矩实际值
		double passBendForceSet;//道次弯辊力设定值
		double passBendForceAct;//道次弯辊力实际值
		double passShiftSet;//道次窜辊设定值
		double passShiftAct;//道次窜辊实际值
		double passVelocity;//道次速度
		int passEntryTemp;//道次入口温度
		int passExitTemp;//道次出口温度
	};

	struct ROLLSchedule{
		PDIData pdiData;//PDI信息
		RollSetup rollSetup;//轧制设定信息
		MillData millDataRM;//粗轧轧机信息
		MillData millDataFM;//精轧轧机信息
		PassData passData[30];//道次信息
	};
}

#endif
