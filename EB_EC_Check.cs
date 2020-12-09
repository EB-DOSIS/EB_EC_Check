//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//++  Eclipde-Script zum monatlichen Eclipse-Check                                               ++
//++  erstes Multi-Class-Script                                                                  ++
//++  mit automatischer Excel-Anbindung                                                          ++
//++  Created by Eyck Blank am 05.10.2020                                                        ++
//++  Version vom 05.10.2020                                                                     ++
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;


using System.Reflection;
using System.Runtime.CompilerServices;
using Microsoft.Office.Core;
// using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.15")]
[assembly: AssemblyFileVersion("1.0.0.15")]
[assembly: AssemblyInformationalVersion("1.15")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{

    //#############################################################################################
    //###              Class DvhExtensions                                                      ###
    //#############################################################################################
    public static class DvhExtensions
    {
		//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		public static DoseValue GetDoseAtVolume
			(this PlanningItem pitem, Structure structure, double volume,
			VolumePresentation volumePresentation, DoseValuePresentation requestedDosePresentation)
		{
			if (pitem is PlanSetup)
			{
				return ((PlanSetup)pitem).GetDoseAtVolume(structure, volume, volumePresentation, requestedDosePresentation);
			}
			else
			{
				if (requestedDosePresentation != DoseValuePresentation.Absolute)
					throw new ApplicationException("Only absolute dose supported for Plan Sums");
				DVHData dvh = pitem.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, volumePresentation, 0.001);
				return DvhExtensions.DoseAtVolume(dvh, volume);
			}
		}

		//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		public static double GetVolumeAtDose
			(this PlanningItem pitem, Structure structure, DoseValue dose, VolumePresentation requestedVolumePresentation)
		{
			if (pitem is PlanSetup)
			{
				return ((PlanSetup)pitem).GetVolumeAtDose(structure, dose, requestedVolumePresentation);
			}
			else
			{
				DVHData dvh = pitem.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, requestedVolumePresentation, 0.001);
				return DvhExtensions.VolumeAtDose(dvh, dose.Dose);
			}
		}

		//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		public static DoseValue DoseAtVolume
			(DVHData dvhData, double volume)
		{
			if (dvhData == null || dvhData.CurveData.Count() == 0)
				return DoseValue.UndefinedDose();
			double absVolume = dvhData.CurveData[0].VolumeUnit == "%" ? volume * dvhData.Volume * 0.01 : volume;
			if (volume < 0.0 || absVolume > dvhData.Volume)
				return DoseValue.UndefinedDose();

			DVHPoint[] hist = dvhData.CurveData;
			for (int i = 0; i < hist.Length; i++)
			{
				if (hist[i].Volume < volume)
					return hist[i].DoseValue;
			}
			return DoseValue.UndefinedDose();
		}

		//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		public static double VolumeAtDose
			(DVHData dvhData, double dose)
		{
			if (dvhData == null)
				return Double.NaN;

			DVHPoint[] hist = dvhData.CurveData;
			int index = (int)(hist.Length * dose / dvhData.MaxDose.Dose);
			if (index < 0 || index > hist.Length)
				return 0.0;//Double.NaN;
			else
				return hist[index].Volume;
		}

	}

    //#############################################################################################
    //###              Class Script                                                             ###
    //#############################################################################################
    public class Script
    {
		//Declaration globaler Variablen

		string Pat = "";
		string PlanPTV = "";
		string PlanID = "";
		string PlanPrescr = "";

		double GD;
		double ED;
		double N;   //GD=ED*N
		double PI;  //prescr.isodose

		PlanningItem SelectedPlanningItem { get; set; }
		StructureSet SelectedStructureSet { get; set; }
		Structure SelectedStructure { get; set; }


		public Script()
        {
			// leer

		}

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
			
			// TODO : Add here the code that is called when the script is launched from Eclipse.
            if (context.Patient != null)
            {
                MessageBox.Show("Patient ID is " + context.Patient.Id);
            }
            else
            {
                MessageBox.Show("No patient is selected");
            }


			//...................................................
			//..   Ab  hier habe ich editiert   05.10.2020     ..
			//..   Auslesen der Feldparameter                  ..
			//...................................................

			Pat = context.Patient.Name.ToString();
			string patID = context.Patient.Id.ToString();

			PlanSetup plan = context.PlanSetup;
			PlanSum psum = context.PlanSumsInScope.FirstOrDefault();
			if (plan == null && psum == null)
				return;

			// Datum zeit
			string Datum = DateTime.Now.ToString("dd.MM.yyyy");
            string Zeit = DateTime.Now.ToString("HH:mm:ss");
			String PlanName = plan.Id;

			//Plan, Structureset
			SelectedPlanningItem = plan != null ? (PlanningItem)plan : (PlanningItem)psum;

			// Plans in plansum can have different structuresets but here we only use structureset to allow chosing one structure
			SelectedStructureSet = plan != null ? plan.StructureSet : psum.PlanSetups.First().StructureSet;
			//
			string PlanID = plan.Id;
			PlanID = PlanID + " (" + plan.ApprovalStatus.ToString() + "/" + plan.HistoryUserName + ")";

			ExternalPlanSetup ExPS = context.ExternalPlanSetup;
			//Beam beam;
		
			double[] MUfield = new double[8];
			int[] controlpointMLC = new int[8];

			string[] sMUfield = new string[8];
			string[] scontrolpointMLC = new string[8];

			// MUs und KPs initielisieren auch für RA-Pläne (2Felser)
			int i = 0;
			for (i = 0; i < 7; i++)
			{
				MUfield[i+1] = 9999;
				controlpointMLC[i+1] = 9999;

				sMUfield[i+1] = MUfield[i+1].ToString("F1");
				scontrolpointMLC[i+1] = controlpointMLC[i+1].ToString("F1");
			}
			i = 0;
			foreach (var beam in ExPS.Beams)
			{ 
				if (!beam.IsSetupField && beam.Meterset.Value.ToString() != "")
				{
					i = i+1;
					MUfield[i] = beam.Meterset.Value;
					controlpointMLC[i] = beam.ControlPoints.Count;

					sMUfield[i] = MUfield[i].ToString("F2");
					scontrolpointMLC[i] =controlpointMLC[i].ToString("F2");
				}

			}

			//...................................................

			PlanPTV = plan.TargetVolumeID;

			//PlanID = PlanID +"/"+ plan.PlanNormalizationValue.ToString("f2");

			GD = plan.TotalDose.Dose;
			ED = plan.DosePerFraction.Dose;
			N = (double)plan.NumberOfFractions.Value;
			PI = plan.TreatmentPercentage * 100;

			// double ED = GD / N;

			string sGD = GD.ToString("F2") + plan.TotalDose.Unit;
			string sED = ED.ToString("F2");
			string sN = N.ToString("F0");
			string sPI = PI.ToString("F2");

			PlanPrescr = PlanPrescr + sED + " x " + sN + " = " + sGD + " (" + sPI + "%)";

			if (SelectedPlanningItem.Dose == null)
				return;


			//------------------------------------------------------------------------------------------------
			//ges Lunge

			Structure Lungeges = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "ges"+" "+"Lungen");
			if (Lungeges == null)
				throw new ApplicationException("no ges Lunge");

			//------------------------------------------------------------------------------------------------
			//Lunge re

			Structure Lungere = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Lunge"+" "+"re");
			if (Lungere == null)
				throw new ApplicationException("no Lunge re");

			//------------------------------------------------------------------------------------------------
			//Lunge li

			Structure Lungeli = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Lunge"+" "+"li");
			if (Lungeli == null)
				throw new ApplicationException("no Lunge li");

			//------------------------------------------------------------------------------------------------
			//GTV-A-BCA-R-74Gy

			Structure GTV74 = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "GTV-A-BCA-R-74Gy");
			if (GTV74 == null)
				throw new ApplicationException("no GTV-A-BCA-R-74Gy");

			//------------------------------------------------------------------------------------------------
			//Herz

			Structure Herz = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Herz");
			if (Herz == null)
				throw new ApplicationException("no Herz");

			//------------------------------------------------------------------------------------------------
			//Implant Ti

			Structure ImplantTi = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Implantat"+" "+"Ti");
			if (ImplantTi == null)
				throw new ApplicationException("no Implant Ti");

			//------------------------------------------------------------------------------------------------
			//Myelon

			Structure Myelon = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Myelon");
			if (Myelon == null)
				throw new ApplicationException("no Myelon");

			//------------------------------------------------------------------------------------------------
			//Oesophagus

			Structure Oesophagus = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Oesophagus");
			if (Oesophagus == null)
				throw new ApplicationException("no Oesophagus");

			//------------------------------------------------------------------------------------------------
			//PTV-A-BCA-R-60Gy

			Structure PTV60 = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "PTV-A-BCA-R-60Gy");
			if (PTV60 == null)
				throw new ApplicationException("no PTV-A-BCA-R-60Gy");

			//------------------------------------------------------------------------------------------------
			//PTV-A-BCA-R-74Gy

			Structure PTV74 = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "PTV-A-BCA-R-74Gy");
			if (PTV74 == null)
				throw new ApplicationException("no PTV-A-BCA-R-74Gy");

			//------------------------------------------------------------------------------------------------
			//Trachea

			Structure Trachea = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Trachea");
			if (Trachea == null)
				throw new ApplicationException("no Trachea");

			//------------------------------------------------------------------------------------------------
			//Z_Implant-Ti

			Structure Z_ImplantTi = SelectedStructureSet.Structures.FirstOrDefault(x => x.Id == "Z_Implantat");
			if (Z_ImplantTi == null)
				throw new ApplicationException("no Z_Implant-Ti");

			//------------------------------------------------------------------------------------------------
			//% des Volumens (PTV)

			double V1 = 0;
			double V2 = 2;
			double V3 = 50;
			double V4 = 98;
			double V5 = 100;

			//................................................................................................          
			//wenn Isodose Absolut gewünscht ist muss statt Relative -Absolute stehen

			DoseValuePresentation dosePres = DoseValuePresentation.Absolute;
			VolumePresentation volPres = VolumePresentation.Relative;

			//................................................................................................
			// Wertübergabe der DVH-Parameter für beide Strukturen
				
				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				DVHData dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Lungeli, dosePres, volPres, s_binWidth);
				double Lungelimean = dvhData2.MeanDose.Dose;
				double Lungelimedian = dvhData2.MedianDose.Dose;
				double Lungelimin = dvhData2.MinDose.Dose;
				double Lungelimax = dvhData2.MaxDose.Dose;

				// für MessageBox
				string LungeliSmin = "Min Dose Lunge li = " + Lungelimin.ToString("F2") + " Gy";
				string LungeliSmax = "Max Dose Lunge li = " + Lungelimax.ToString("F2") + " Gy";
				string LungeliSmean = "Mean Dose Lunge li = " + Lungelimean.ToString("F2") + " Gy";
				string LungeliSmedian = "Median Dose Lunge li = " + Lungelimedian.ToString("F2") + " Gy";
				// für Excel
				string sLungelimin = Lungelimin.ToString("F2");
				string sLungelimax = Lungelimax.ToString("F2");
				string sLungelimean = Lungelimean.ToString("F2");
				string sLungelimedian = Lungelimedian.ToString("F2");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//2) Lunge re

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Lungere, dosePres, volPres, s_binWidth);
				double Lungeremean = dvhData2.MeanDose.Dose;
				double Lungeremedian = dvhData2.MedianDose.Dose;
				double Lungeremin = dvhData2.MinDose.Dose;
				double Lungeremax = dvhData2.MaxDose.Dose;

				string LungereSmin = "Min Dose Lunge li = " + Lungeremin.ToString("F3") + " Gy";
				string LungereSmax = "Max Dose Lunge li = " + Lungeremax.ToString("F3") + " Gy";
				string LungereSmean = "Mean Dose Lunge li = " + Lungeremean.ToString("F3") + " Gy";
				string LungereSmedian = "Median Dose Lunge li = " + Lungeremedian.ToString("F3") + " Gy";

				string sLungeremin = Lungeremin.ToString("F3");
				string sLungeremax = Lungeremax.ToString("F3");
				string sLungeremean = Lungeremean.ToString("F3");
				string sLungeremedian = Lungeremedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//3) Lunge ges

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Lungeges, dosePres, volPres, s_binWidth);
				double Lungegesmean = dvhData2.MeanDose.Dose;
				double Lungegesmedian = dvhData2.MedianDose.Dose;
				double Lungegesmin = dvhData2.MinDose.Dose;
				double Lungegesmax = dvhData2.MaxDose.Dose;

				string LungegesSmin = "Min Dose Lunge ges = " + Lungegesmin.ToString("F3") + " Gy";
				string LungegesSmax = "Max Dose Lunge ges = " + Lungegesmax.ToString("F3") + " Gy";
				string LungegesSmean = "Mean Dose Lunge ges = " + Lungegesmean.ToString("F3") + " Gy";
				string LungegesSmedian = "Median Dose Lunge ges = " + Lungegesmedian.ToString("F3") + " Gy";

				string sLungegesmin = Lungegesmin.ToString("F3");
				string sLungegesmax = Lungegesmax.ToString("F3");
				string sLungegesmean = Lungegesmean.ToString("F3");
				string sLungegesmedian = Lungegesmedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//4) GTV-A-BCA-R-74Gy

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(GTV74, dosePres, volPres, s_binWidth);
				double GTV74mean = dvhData2.MeanDose.Dose;
				double GTV74median = dvhData2.MedianDose.Dose;
				double GTV74min = dvhData2.MinDose.Dose;
				double GTV74max = dvhData2.MaxDose.Dose;

				string GTV74Smin = "Min Dose GTV-74 = " + GTV74min.ToString("F3") + " Gy";
				string GTV74Smax = "Max Dose GTV-74 = " + GTV74max.ToString("F3") + " Gy";
				string GTV74Smean = "Mean Dose GTV-74 = " + GTV74mean.ToString("F3") + " Gy";
				string GTV74Smedian = "Median Dose GTV-74 = " + GTV74median.ToString("F3") + " Gy";

				string sGTV74min = GTV74min.ToString("F3");
				string sGTV74max = GTV74max.ToString("F3");
				string sGTV74mean = GTV74mean.ToString("F3");
				string sGTV74median = GTV74median.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//5) Herz

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Herz, dosePres, volPres, s_binWidth);
				double Herzmean = dvhData2.MeanDose.Dose;
				double Herzmedian = dvhData2.MedianDose.Dose;
				double Herzmin = dvhData2.MinDose.Dose;
				double Herzmax = dvhData2.MaxDose.Dose;

				string HerzSmin = "Min Dose Herz = " + Herzmin.ToString("F3") + " Gy";
				string HerzSmax = "Max Dose Herz = " + Herzmax.ToString("F3") + " Gy";
				string HerzSmean = "Mean Dose Herz = " + Herzmean.ToString("F3") + " Gy";
				string HerzSmedian = "Median Dose Herz = " + Herzmedian.ToString("F3") + " Gy";

				string sHerzmin = Herzmin.ToString("F3");
				string sHerzmax = Herzmax.ToString("F3");
				string sHerzmean = Herzmean.ToString("F3");
				string sHerzmedian = Herzmedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//6) Implant Ti

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(ImplantTi, dosePres, volPres, s_binWidth);
				double ImplantTimean = dvhData2.MeanDose.Dose;
				double ImplantTimedian = dvhData2.MedianDose.Dose;
				double ImplantTimin = dvhData2.MinDose.Dose;
				double ImplantTimax = dvhData2.MaxDose.Dose;

				string ImplantTiSmin = "Min Dose ImplantTi = " + ImplantTimin.ToString("F3") + " Gy";
				string ImplantTiSmax = "Max Dose ImplantTi = " + ImplantTimax.ToString("F3") + " Gy";
				string ImplantTiSmean = "Mean Dose ImplantTi = " + ImplantTimean.ToString("F3") + " Gy";
				string ImplantTiSmedian = "Median Dose ImplantTi = " + ImplantTimedian.ToString("F3") + " Gy";

				string sImplantTimin = ImplantTimin.ToString("F3");
				string sImplantTimax = ImplantTimax.ToString("F3");
				string sImplantTimean = ImplantTimean.ToString("F3");
				string sImplantTimedian = ImplantTimedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//7) Myelon

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Myelon, dosePres, volPres, s_binWidth);
				double Myelonmean = dvhData2.MeanDose.Dose;
				double Myelonmedian = dvhData2.MedianDose.Dose;
				double Myelonmin = dvhData2.MinDose.Dose;
				double Myelonmax = dvhData2.MaxDose.Dose;

				string MyelonSmin = "Min Dose Myelon = " + Myelonmin.ToString("F3") + " Gy";
				string MyelonSmax = "Max Dose Myelon = " + Myelonmax.ToString("F3") + " Gy";
				string MyelonSmean = "Mean Dose Myelon = " + Myelonmean.ToString("F3") + " Gy";
				string MyelonSmedian = "Median Dose Myelon = " + Myelonmedian.ToString("F3") + " Gy";

				string sMyelonmin = Myelonmin.ToString("F3");
				string sMyelonmax = Myelonmax.ToString("F3");
				string sMyelonmean = Myelonmean.ToString("F3");
				string sMyelonmedian = Myelonmedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//8) Oesophagus

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Oesophagus, dosePres, volPres, s_binWidth);
				double Oesomean = dvhData2.MeanDose.Dose;
				double Oesomedian = dvhData2.MedianDose.Dose;
				double Oesomin = dvhData2.MinDose.Dose;
				double Oesomax = dvhData2.MaxDose.Dose;

				string OesoSmin = "Min Dose Oesophagus = " + Oesomin.ToString("F3") + " Gy";
				string OesoSmax = "Max Dose Oesophagus = " + Oesomax.ToString("F3") + " Gy";
				string OesoSmean = "Mean Dose Oesophagus = " + Oesomean.ToString("F3") + " Gy";
				string OesoSmedian = "Median Dose Oesophagus = " + Oesomedian.ToString("F3") + " Gy";

				string sOesomin = Oesomin.ToString("F3");
				string sOesomax = Oesomax.ToString("F3");
				string sOesomean = Oesomean.ToString("F3");
				string sOesomedian = Oesomedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//9) PTV-A-BCA-R-60Gy

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(PTV60, dosePres, volPres, s_binWidth);
				double PTV60mean = dvhData2.MeanDose.Dose;
				double PTV60median = dvhData2.MedianDose.Dose;
				double PTV60min = dvhData2.MinDose.Dose;
				double PTV60max = dvhData2.MaxDose.Dose;

				string PTV60Smin = "Min Dose PTV60 = " + PTV60min.ToString("F3") + " Gy";
				string PTV60Smax = "Max Dose PTV60 = " + PTV60max.ToString("F3") + " Gy";
				string PTV60Smean = "Mean Dose PTV60 = " + PTV60mean.ToString("F3") + " Gy";
				string PTV60Smedian = "Median Dose PTV60 = " + PTV60median.ToString("F3") + " Gy";
			
				string sPTV60min = PTV60min.ToString("F3");
				string sPTV60max = PTV60max.ToString("F3");
				string sPTV60mean = PTV60mean.ToString("F3");
				string sPTV60median = PTV60median.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//10) PTV-A-BCA-R-74Gy

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(PTV74, dosePres, volPres, s_binWidth);
				double PTV74mean = dvhData2.MeanDose.Dose;
				double PTV74median = dvhData2.MedianDose.Dose;
				double PTV74min = dvhData2.MinDose.Dose;
				double PTV74max = dvhData2.MaxDose.Dose;

				string PTV74Smin = "Min Dose PTV74 = " + PTV74min.ToString("F3") + " Gy";
				string PTV74Smax = "Max Dose PTV74 = " + PTV74max.ToString("F3") + " Gy";
				string PTV74Smean = "Mean Dose PTV74 = " + PTV74mean.ToString("F3") + " Gy";
				string PTV74Smedian = "Median Dose PTV74 = " + PTV74median.ToString("F3") + " Gy";

				string sPTV74min = PTV74min.ToString("F3");
				string sPTV74max = PTV74max.ToString("F3");
				string sPTV74mean = PTV74mean.ToString("F3");
				string sPTV74median = PTV74median.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//11) Trachea

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(Trachea, dosePres, volPres, s_binWidth);
				double Tracheamean = dvhData2.MeanDose.Dose;
				double Tracheamedian = dvhData2.MedianDose.Dose;
				double Tracheamin = dvhData2.MinDose.Dose;
				double Tracheamax = dvhData2.MaxDose.Dose;

				string TracheaSmin = "Min Dose Trachea = " + Tracheamin.ToString("F3") + " Gy";
				string TracheaSmax = "Max Dose Trachea = " + Tracheamax.ToString("F3") + " Gy";
				string TracheaSmean = "Mean Dose Trachea = " + Tracheamean.ToString("F3") + " Gy";
				string TracheaSmedian = "Median Dose Trachea = " + Tracheamedian.ToString("F3") + " Gy";

				string sTracheamin = Tracheamin.ToString("F3");
				string sTracheamax = Tracheamax.ToString("F3");
				string sTracheamean = Tracheamean.ToString("F3");
				string sTracheamedian = Tracheamedian.ToString("F3");

				//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
				//12) Z_Implant-Ti

				dosePres = DoseValuePresentation.Absolute;
				volPres = VolumePresentation.Relative;

				dvhData2 = SelectedPlanningItem.GetDVHCumulativeData(ImplantTi, dosePres, volPres, s_binWidth);
				double Z_ImplTimean = dvhData2.MeanDose.Dose;
				double Z_ImplTimedian = dvhData2.MedianDose.Dose;
				double Z_ImplTimin = dvhData2.MinDose.Dose;
				double Z_ImplTimax = dvhData2.MaxDose.Dose;

				string Z_ImplTiSmin = "Min Dose Z_ImplantTi = " + Z_ImplTimin.ToString("F3") + " Gy";
				string Z_ImplTiSmax = "Max Dose Z_ImplantTi = " + Z_ImplTimax.ToString("F3") + " Gy";
				string Z_ImplTiSmean = "Mean Dose Z_ImplantTi = " + Z_ImplTimean.ToString("F3") + " Gy";
				string Z_ImplTiSmedian = "Median Dose Z_ImplantTi = " + Z_ImplTimedian.ToString("F3") + " Gy";

				string sZ_ImplTimin = Z_ImplTimin.ToString("F3");
				string sZ_ImplTimax = Z_ImplTimax.ToString("F3");
				string sZ_ImplTimean = Z_ImplTimean.ToString("F3");
				string sZ_ImplTimedian = Z_ImplTimedian.ToString("F3");


				// window.Title = Pat;


				if(MessageBox.Show(Datum  + " " + Zeit + "\r\n" + PlanName + 
					"\r\n" + "\r\n" + "ok = Abspeichern"  + "\r\n" + "cancel = nicht speichern", 
					"EB_EC_Check", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
				{
					// Bei MessageBox = ok dann schreibe in Excel
					UpdateExcel(Datum, Zeit, patID, PlanName, 
						sMUfield[1], scontrolpointMLC[1],   sMUfield[2], scontrolpointMLC[2],   sMUfield[3], scontrolpointMLC[3],   sMUfield[4], scontrolpointMLC[4], 
						sLungegesmin,  sLungegesmax,  sLungegesmean,  sLungegesmedian, 
						sGTV74min,     sGTV74max,     sGTV74mean,     sGTV74median, 
						sHerzmin,      sHerzmax,      sHerzmean,      sHerzmedian, 
						sImplantTimin, sImplantTimax, sImplantTimean, sImplantTimedian, 
						sLungelimin,   sLungelimax,   sLungelimean,   sLungelimedian, 
						sLungeremin,   sLungeremax,   sLungeremean,   sLungeremedian, 
						sMyelonmin,    sMyelonmax,    sMyelonmean,    sMyelonmedian, 
						sOesomin,      sOesomax,      sOesomean,      sOesomedian, 
						sPTV60min,     sPTV60max,     sPTV60mean,     sPTV60median, 
						sPTV74min,     sPTV74max,     sPTV74mean,     sPTV74median, 
						sTracheamin,   sTracheamax,   sTracheamean,   sTracheamedian, 
						sZ_ImplTimin,  sZ_ImplTimax,  sZ_ImplTimean,  sZ_ImplTimedian);
					
				}
				else
				{  
					// nothing // Close();
				}  

			
		}
		static double s_binWidth = 0.001;
		
		//##########################################################################################################
		
        private void UpdateExcel(string xlsDatum, string xlsZeit, string xlsPatID, string xlsPlanName, 
            string F1MU, string F1KP,   string F2MU, string F2KP,   string F3MU, string F3KP,   string F4MU, string F4KP, 
			string gesLuMin,   string gesLuMax,   string gesLuMean,   string gesLuMedian, 
			string GTV74Min,   string GTV74Max,   string GTV74Mean,   string GTV74Median,
			string HerzMin,    string HerzMax,    string HerzMean,    string HerzMedian,
			string ImplTiMin,  string ImplTiMax,  string ImplTiMean,  string ImplTiMedian,
			string LuliMin,    string LuliMax,    string LuliMean,    string LuliMedian,
			string LureMin,    string LureMax,    string LureMean,    string LureMedian,
			string MyolMin,    string MyolMax,    string MyolMean,    string MyolMedian,
			string OesoMin,    string OesoMax,    string OesoMean,    string OesoMedian,
			string PTV60Min,   string PTV60Max,   string PTV60Mean,   string PTV60Median,
			string PTV74Min,   string PTV74Max,   string PTV74Mean,   string PTV74Median,
			string TrachMin,   string TrachMax,   string TrachMean,   string TrachMedian,
			string ZimplMin,   string ZimplMax,   string ZimplMean,   string ZimplMedian)
		
		{

			string Plan2258imX06ac = "IM-2258-X6";
			string Plan2258imX15ac = "IM-2258-X15";
			string Plan2258raX06ac = "RA-2258-X6";
			string Plan2258raX15ac = "RA-2258-X15";

			string Plan4160imX06ac = "IM-4160-X6";
			string Plan4160imX15ac = "IM-4160-X15";
			string Plan4160raX06ac = "RA-4160-X6";
			string Plan4160raX15ac = "RA-4160-X15";

			string Plan4434imX06ac = "IM-4434-X6";
			string Plan4434imX15ac = "IM-4434-X15";
			string Plan4434raX06ac = "RA-4434-X6";
			string Plan4434raX15ac = "RA-4434-X15";

			string Plan4160imX06aa = "IM-4160-X6-AA";
			string Plan4160imX15aa = "IM-4160-X15-A";
			string Plan4160raX06aa = "RA-4160-X6-AA";
			string Plan4160raX15aa = "RA-4160-X15-A";

			string Plan4434imX06aa = "IM-4434-X6-AA";
			string Plan4434imX15aa = "IM-4434-X15-A";
			string Plan4434raX06aa = "RA-4434-X6-AA";
			string Plan4434raX15aa = "RA-4434-X15-A";

			string xlsDatei2258ac = "Q:/Eclipse/QA_EC_A15_xls/EC_Ac_2258_NP_A15.xlsx";
			string xlsDatei4160ac = "Q:/Eclipse/QA_EC_A15_xls/EC_Ac_4160_NP_A15.xlsx";
			string xlsDatei4434ac = "Q:/Eclipse/QA_EC_A15_xls/EC_Ac_4434_NP_A15.xlsx";
			string xlsDatei4160aa = "Q:/Eclipse/QA_EC_A15_xls/EC_AAA_4160_NP_A15.xlsx";
			string xlsDatei4434aa = "Q:/Eclipse/QA_EC_A15_xls/EC_AAA_4434_NP_A15.xlsx";
			string xlsDatei = " ";

			string xlsAB2258imX06 = "2258_IM_X6";
			string xlsAB2258imX15 = "2258_IM_X15";
			string xlsAB2258raX06 = "2258_RA_X6";
			string xlsAB2258raX15 = "2258_RA_X15";

			string xlsAB4160imX06 = "4160_IM_X6";
			string xlsAB4160imX15 = "4160_IM_X15";
			string xlsAB4160raX06 = "4160_RA_X6";
			string xlsAB4160raX15 = "4160_RA_X15";

			string xlsAB4434imX06 = "4434_IM_X6";
			string xlsAB4434imX15 = "4434_IM_X15";
			string xlsAB4434raX06 = "4434_RA_X6";
			string xlsAB4434raX15 = "4434_RA_X15";

			string xlsAB = " ";
			



            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;


			// Automatik fürs richtige Abspeichern
			// .....................................
			// 2258
			if (xlsPlanName == Plan2258imX06ac)
			{
				xlsDatei = xlsDatei2258ac;
				xlsAB = xlsAB2258imX06;
			
			};
			if (xlsPlanName == Plan2258imX15ac)
			{
				xlsDatei =  xlsDatei2258ac;
				xlsAB = xlsAB2258imX15;

			};
			if (xlsPlanName == Plan2258raX06ac)
			{
				xlsDatei =  xlsDatei2258ac;
				xlsAB = xlsAB2258raX06;

			};
			if (xlsPlanName == Plan2258raX15ac)
			{
				xlsDatei =  xlsDatei2258ac;
				xlsAB = xlsAB2258raX15;
			
			};
			// 4160
			if (xlsPlanName == Plan4160imX06ac)
			{
				xlsDatei =  xlsDatei4160ac;
				xlsAB = xlsAB4160imX06;
			
			};
			if (xlsPlanName == Plan4160imX15ac)
			{
				xlsDatei =  xlsDatei4160ac;
				xlsAB = xlsAB4160imX15;
			
			};
			if (xlsPlanName == Plan4160raX06ac)
			{
				xlsDatei =  xlsDatei4160ac;
				xlsAB = xlsAB4160raX06;
			
			};
			if (xlsPlanName == Plan4160raX15ac)
			{
				xlsDatei =  xlsDatei4160ac;
				xlsAB = xlsAB4160raX15;
			
			};
			// 4434
			if (xlsPlanName == Plan4434imX06ac)
			{
				xlsDatei =  xlsDatei4434ac;
				xlsAB = xlsAB4434imX06;
			
			};
			if (xlsPlanName == Plan4434imX15ac)
			{
				xlsDatei =  xlsDatei4434ac;
				xlsAB = xlsAB4434imX15;
			
			};
			if (xlsPlanName == Plan4434raX06ac)
			{
				xlsDatei =  xlsDatei4434ac;
				xlsAB = xlsAB4434raX06;
			
			};
			if (xlsPlanName == Plan4434raX15ac)
			{
				xlsDatei =  xlsDatei4434ac;
				xlsAB = xlsAB4434raX15;

			};
			// 4160 AAA
			if (xlsPlanName == Plan4160imX06aa)
			{
				xlsDatei = xlsDatei4160aa;
				xlsAB = xlsAB4160imX06;

			};
			if (xlsPlanName == Plan4160imX15aa)
			{
				xlsDatei = xlsDatei4160aa;
				xlsAB = xlsAB4160imX15;

			};
			if (xlsPlanName == Plan4160raX06aa)
			{
				xlsDatei = xlsDatei4160aa;
				xlsAB = xlsAB4160raX06;

			};
			if (xlsPlanName == Plan4160raX15aa)
			{
				xlsDatei = xlsDatei4160aa;
				xlsAB = xlsAB4160raX15;

			};
			// 4434 AAA
			if (xlsPlanName == Plan4434imX06aa)
			{
				xlsDatei =  xlsDatei4434aa;
				xlsAB = xlsAB4434imX06;
			};
			if (xlsPlanName == Plan4434imX15aa)
			{
				xlsDatei =  xlsDatei4434aa;
				xlsAB = xlsAB4434imX15;
			};
			if (xlsPlanName == Plan4434raX06aa)
			{
				xlsDatei =  xlsDatei4434aa;
				xlsAB = xlsAB4434raX06;
			};
			if (xlsPlanName == Plan4434raX15aa)
			{
				xlsDatei =  xlsDatei4434aa;
				xlsAB = xlsAB4434raX15;
			};


            try
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open(xlsDatei);
                oSheet = String.IsNullOrEmpty(xlsAB) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[xlsAB];

                string sReihe = oSheet.Cells[2, 66].Value == null ? "-" : oSheet.Cells[2, 66].Value.ToString();
                int iReihe = Convert.ToInt32(sReihe);
                iReihe = iReihe + 1;

                oSheet.Cells[iReihe, 1] = xlsDatum;
                oSheet.Cells[iReihe, 2] = xlsZeit;

				oSheet.Cells[iReihe, 3] = xlsPatID;
				oSheet.Cells[iReihe, 4] = xlsPlanName;

				oSheet.Cells[iReihe, 8] = F1MU;
                oSheet.Cells[iReihe, 9] = F1KP;

                oSheet.Cells[iReihe, 10] = F2MU;
                oSheet.Cells[iReihe, 11] = F2KP;

                oSheet.Cells[iReihe, 12] = F3MU;
                oSheet.Cells[iReihe, 13] = F3KP;

                oSheet.Cells[iReihe, 14] = F4MU;
				oSheet.Cells[iReihe, 15] = F4KP;


                oSheet.Cells[iReihe, 17] = gesLuMin;
                oSheet.Cells[iReihe, 18] = gesLuMax;
                oSheet.Cells[iReihe, 19] = gesLuMean;
                oSheet.Cells[iReihe, 20] = gesLuMedian;
				
				oSheet.Cells[iReihe, 21] = GTV74Min;
                oSheet.Cells[iReihe, 22] = GTV74Max;
                oSheet.Cells[iReihe, 23] = GTV74Mean;
                oSheet.Cells[iReihe, 24] = GTV74Median;

				oSheet.Cells[iReihe, 25] = HerzMin;
                oSheet.Cells[iReihe, 26] = HerzMax;
                oSheet.Cells[iReihe, 27] = HerzMean;
                oSheet.Cells[iReihe, 28] = HerzMedian;

				oSheet.Cells[iReihe, 29] = ImplTiMin;
                oSheet.Cells[iReihe, 30] = ImplTiMax;
                oSheet.Cells[iReihe, 31] = ImplTiMean;
                oSheet.Cells[iReihe, 32] = ImplTiMedian;

				oSheet.Cells[iReihe, 33] = LuliMin;
                oSheet.Cells[iReihe, 34] = LuliMax;
                oSheet.Cells[iReihe, 35] = LuliMean;
                oSheet.Cells[iReihe, 36] = LuliMedian;

				oSheet.Cells[iReihe, 37] = LureMin;
                oSheet.Cells[iReihe, 38] = LureMax;
                oSheet.Cells[iReihe, 39] = LureMean;
                oSheet.Cells[iReihe, 40] = LureMedian;

				oSheet.Cells[iReihe, 41] = MyolMin;
                oSheet.Cells[iReihe, 42] = MyolMax;
                oSheet.Cells[iReihe, 43] = MyolMean;
                oSheet.Cells[iReihe, 44] = MyolMedian;

				oSheet.Cells[iReihe, 45] = OesoMin;
                oSheet.Cells[iReihe, 46] = OesoMax;
                oSheet.Cells[iReihe, 47] = OesoMean;
                oSheet.Cells[iReihe, 48] = OesoMedian;

				oSheet.Cells[iReihe, 49] = PTV60Min;
                oSheet.Cells[iReihe, 50] = PTV60Max;
                oSheet.Cells[iReihe, 51] = PTV60Mean;
                oSheet.Cells[iReihe, 52] = PTV60Median;

				oSheet.Cells[iReihe, 53] = PTV74Min;
                oSheet.Cells[iReihe, 54] = PTV74Max;
                oSheet.Cells[iReihe, 55] = PTV74Mean;
                oSheet.Cells[iReihe, 56] = PTV74Median;

				oSheet.Cells[iReihe, 57] = TrachMin;
                oSheet.Cells[iReihe, 58] = TrachMax;
                oSheet.Cells[iReihe, 59] = TrachMean;
                oSheet.Cells[iReihe, 60] = TrachMedian;
								
				oSheet.Cells[iReihe, 61] = ZimplMin;
                oSheet.Cells[iReihe, 62] = ZimplMax;
                oSheet.Cells[iReihe, 63] = ZimplMean;
                oSheet.Cells[iReihe, 64] = ZimplMedian;
								
                sReihe = Convert.ToString(iReihe);

                oSheet.Cells[2, 66] = sReihe;

                oWB.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null) 
                {
                  oWB.Close(true, null, null);
                  oXL.Quit();   
                }
                    
            }
            
            MessageBox.Show("Done");

        }


	}

}
