#ifndef ROLLSCHEDULEDEFINITION_H
#define ROLLSCHEDULEDEFINITION_H

namespace PDFCreater{

	//PDI��Ϣ
	struct PDIData{
		char steelGrade[256];//������
		double slabWidthL3;//�������-L3[mm]
		double slabWidthAct;//�������-ʵ��[mm]
		double slabThickL3;//�������-L3[mm]
		double slabThickAct;//�������-ʵ��[mm]
		double slabLengthL3;//��������-L3[mm]
		double slabLengthAct;//��������-ʵ��[mm]
		double slabWeightL3;//��������-L3
		double slabWeightAct;//��������-ʵ��
		int furnaceNo;//����¯��
		int furnaceCol;//����¯��

		double plateThickSet;//��ƷĿ����[mm]
		double plateThickAct;//��Ʒʵ�ʺ��[mm]
		double plateWidthSet;//��ƷĿ����[mm]
		double plateWidthAct;//��Ʒʵ�ʿ��[mm]
		double plateLengthSet;//��ƷĿ�곤��[mm]
		double plateLengthAct;//��Ʒʵ�ʳ���[mm]
		int finalTemp;//�����¶�
	};

	//�����趨��Ϣ
	struct RollSetup{
		//����ģʽ
		enum Enum_CtrlMode{
			NOCTRL,//�ǿ���
			ONECTRL,//�������
			TWOCTRL,//�������
			THRCTRL,//�������
			CHAINCTRL//��ʽ����
		};
		//ת��ģʽ
		enum Enum_TurnMode{
			TM_V,//����
			TM_V_H,//��-����
			TM_V_H_V,//��-��-����
			TM_H,//����
			TM_H_V,//��-����
			TM_H_V_H//��-��-����
		};
		int ctrlMode;//����ģʽ��0ΪRM,1ΪFM
		int totalPassNum;//�ܵ�����
		int turnMode;//ת��ģʽ
		int turnPassNum;//ת�ֵ���
		double turnThick;//ת�ֺ��[mm]
		double controlThick;//�������[mm]
	};

	//������Ϣ
	struct RollData{
		char RollID[256];//�������
		char RollMaterial[256];//��������
		char RollStartTime[256];//����ֱ��[mm]
		double RollDiameter;//�ϻ�ʱ��
	};

	//����������Ϣ
	struct MillData{
		RollData upWorkRoll;//�Ϲ�������Ϣ
		RollData downWorkRoll;//�¹�������Ϣ
		RollData upBackupRoll;//��֧�Ź���Ϣ
		RollData downBackupRoll;//��֧�Ź���
	};

	//������Ϣ
	struct PassData{
		int millType;//��������(RM/FM)
		int passNo;//���κ�
		double passGapSet;//���ι����趨ֵ
		double passGapAct;//���ι���ʵ��ֵ
		double passThickSet;//���κ���趨ֵ
		double passThickAct;//���κ��ʵ��ֵ
		double passWidthSet;//���ο���趨ֵ
		double passWidthAct;//���ο��ʵ��ֵ
		double passForceSet;//�����������趨ֵ
		double passForceAct;//����������ʵ��ֵ
		double passTorqueSet;//���������趨ֵ
		double passTorqueAct;//��������ʵ��ֵ
		double passBendForceSet;//����������趨ֵ
		double passBendForceAct;//���������ʵ��ֵ
		double passShiftSet;//���δܹ��趨ֵ
		double passShiftAct;//���δܹ�ʵ��ֵ
		double passVelocity;//�����ٶ�
		int passEntryTemp;//��������¶�
		int passExitTemp;//���γ����¶�
	};

	struct ROLLSchedule{
		PDIData pdiData;//PDI��Ϣ
		RollSetup rollSetup;//�����趨��Ϣ
		MillData millDataRM;//����������Ϣ
		MillData millDataFM;//����������Ϣ
		PassData passData[30];//������Ϣ
	};
}

#endif
