use strict;
use File::Find;
use Cwd;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
my @ARR;my $i=1;$ARR[0]='';
my $Book1;    my $raw="Raw";
my @A=(0..100);
my @ARR;
my @row2=("Time","PP2 (0)","pp0 (0)","PKG (0)","GTCV (0)","IACV (0)","PP0_temp (0)","PP1_temp (0)","IA_C6/C7_res (0)","IA_C6/C7_res (1)","IA_C6/C7_res (2)","IA_C6/C7_res (3)","IA_C3_res (0)","IA_C3_res (1)","IA_C3_res (2)","IA_C3_res (3)","GT_DRAM_BW (0)","IA_DRAM_BW (0)","IO_DRAM_BW (0)","Time (seconds)","GT Power","IA Power","PKG Power","GT Freq","IA Freq","C6 - C7","C3","C0","C6 - C7","C3","C0","C6 - C7","C3","C0","C6 - C7","C3","C0","GT_DRAM_BW (0)","IA_DRAM_BW (0)","IO_DRAM_BW (0)","Total DRAM BW","Time (seconds)","Gfx Power","GFX Vcc","IA Power","IA Vcc","Package Power","Gfx Vcc","Measured GT Power","Reported GT Power","GT Temperature","GT Leakage","IA Vcc","Measured IA Power","Reported IA Power","IA Temperature","IA + Ring/LLC Leakage","Ring/LLC Leakage","IA C6/C7 Residency","IA core Only Leakage","Measured Package Power","Reported Package Power","GT Cdyn","IA + Ring Cdyn","GT+IA Cdyn","GT Max Cdyn","GT App Ratio","IA App Ratio","IA Core0","IA Core1","IA Core2","IA Core3","IA 2 core C0 Avg","IA Core0","IA Core1","IA Core2","IA Core3","IA 2 core C6 Avg","IA Core0","IA Core1","IA Core2","IA Core3","IA 2 core C3 Avg","GT","IA","IO Bandwidth","Total","IA LLC BW","GT LLC BW","Total LLC BW","SA Power","SA Vcc","VDDQ Power","VDDQ Vcc","VCCP (Uncore) Power","VCCP (Uncore) Vcc","SFR Power","SFR Vcc","SA ","VDDQ ","VCCP ","SFR","IA Total C0 ( Sum of All core C0 residency)");
$Win32::OLE::Warn = 0;                                # die on errors...
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit');
	$Excel->{DisplayAlerts}=0;  
foreach my $c ("A".."Z")
{
    $ARR[$i]=$c;
    $i++;
}
sub coloumns
{
    foreach my $k ("A".."Z")
        {
            $ARR[$i]=$_[0].$k;
            $i++;
        }
}
sub Last
{
    my $LR = $_[0]->UsedRange->Find({What=>"*",SearchDirection=>xlPrevious,SearchOrder=>xlByRows})->{Row};
    my $LC = $_[0]->UsedRange->Find({What=>"*",SearchDirection=>xlPrevious,SearchOrder=>xlByColumns})->{Column};
    my @ar=($LR,$LC);
    return @ar;
}

sub Active
{
	my $na = $_[0]->{Name};
    #print "Activating --> $na\n";
	$_[1] =$_[0]->Activate();
	#$_[0]->Worksheets($name)->Activate();
}
sub Merger
{
	#print "$_[1]\n";
	$_[0]->Range("$_[1]")->Merge;
	$_[0]->Range("$_[1]")->{Value}="$_[2]";
	$_[0]->Range("$_[1]")->{HorizontalAlignment}= xlHAlignCenter;
	$_[0]->Range("$_[1]")->{Borders}->{Weight}=xlMedium;
	# $_[0]->Range("$_[1]")->{Borders(eval(xlEdgeTop))}->{Weight}=xlMedium;
	# $_[0]->Range("$_[1]")->{Borders(eval(xlEdgeBottom))}->{Weight}=xlMedium;
	# $_[0]->Range("$_[1]")->{Borders(eval(xlEdgeRight))}->{Weight}=xlMedium;
	# # my @edges = qw (xlInsideHorizontal xlInsideVertical);
	# my $range = "$_[1]"; 
	# foreach my $edge (@edges)
		# {
		  # with (my $Borders = $_[0]->Range($range)->Borders(eval($edge)), 
				  # LineStyle =>xlContinuous,
				  # Weight => xlMedium ,
				  # ColorIndex => 1);
		# }
}
sub Row_2
{
	my $sheet=$_[0];
	my @last=Last($_[0]);
	my $i=1;
	foreach my $k(@row2)
	{
		$sheet->Range("$ARR[$i]2")->{Value}=$k;
		$sheet->Range("$ARR[$i]2")->{Borders}->{Weight}=xlMedium;
		$sheet->Range("$ARR[$i]2")->{Orientation}=90;
		$i++;
	}
}
		
	
foreach my $cc ("A".."Z")
        {
            coloumns($cc);
        }
		
sub wanted
{	
	if($_=~/EMON/)
	{
		my @emon_arr = glob("*EMON*");
		my @daq_arr = glob("*math*");
		
		my $Emon=$emon_arr[0];
		my $Daq=$daq_arr[0];
		my $path = getcwd();
		if ($path=~s/\//\\\\/g)
		{}
		my $abs_path_d=$path.'\\'.$Daq;
		my $abs_path_e=$path.'\\'.$Emon;
	    my $name;
        if ($Daq=~/(\S+)_math/)
        {   
            $name = $1;
        }
		my $Emon_Book = $Excel->Workbooks->Open("$abs_path_e");
		my $Emon_Sheet = $Emon_Book->Worksheets(1);
		my @e_last = Last($Emon_Sheet);
		$Excel->{SheetsInNewWorkBook}=1;
		my $PP_Book=$Excel->Workbooks->Add();my $temp_str = $PP_Book;
									##################### AVERAGE SHEET ###########################
		$PP_Book->Worksheets->Add;
		my $pp_ave_sheet=$PP_Book->Worksheets(1);
		$pp_ave_sheet->{Name}="Experiment_Summary";
									###################### EMON SHEET #############################
		$PP_Book->Worksheets->Add;
		my $pp_emon_sheet=$PP_Book->Worksheets(1);
		$pp_emon_sheet->{Name}="EMON";

									####################### DAQ SHEET #############################
		$PP_Book->Worksheets->Add;
		my $pp_daq_sheet=$PP_Book->Worksheets(1);
		$pp_daq_sheet->{Name}="NI_POWER";
	
		$PP_Book->Worksheets->Add;
		my $pp_lkg_sheet=$PP_Book->Worksheets(1);
		$pp_lkg_sheet->{Name}="Leakage_Lookup";
									##################### Post Process SHEET #####################
		$PP_Book->Worksheets->Add;
		my $pp_sheet=$PP_Book->Worksheets(1);
		$pp_sheet->{Name}="Data";
		
		foreach my $k(1..5)
		{
			my $name= $PP_Book->Worksheets($k)->{Name};
            #print "$k-->$name\n";
		}
		###################################REFERENCE USAGE --------#Active($pp_emon_sheet,$PP_Book);


		#Active($Emon_Book,$Emon_Sheet);
		$Emon_Sheet->Range("A1:$ARR[$e_last[1]]$e_last[0]")->Copy();
		Active($pp_emon_sheet,$PP_Book);
		$pp_emon_sheet->Range("A1:$ARR[$e_last[1]]$e_last[0]")->PasteSpecial();
		$Emon_Book->Close();
		

		## DAQ DATA
		
		my $Daq_Book = $Excel->Workbooks->Open("$abs_path_d");
		my $Daq_Sheet = $Daq_Book->Worksheets(1);
		my @d_last = Last($Daq_Sheet);
	   
		#Active($Daq_Sheet, $Daq_Book);
		$Daq_Sheet->Range("A1:$ARR[$d_last[1]-1]$d_last[0]")->Copy();
		Active($pp_daq_sheet,$PP_Book);
		$pp_daq_sheet->Range("A1:$ARR[$d_last[1]-1]$d_last[0]")->PasteSpecial();
		$Daq_Book->Close();
		#<STDIN>;

		####LEAKAGE SHEET
		my $Lkg_Book = $Excel->Workbooks->Open("c:\\Users\\Public\\Leakage_Sheet_Part2.xlsx");
		my $Lkg_Sheet = $Lkg_Book->Worksheets(1);
		my @lkg_last = Last($Lkg_Sheet);
	   
		
		$Lkg_Sheet->Range("A1:AT57")->Copy();
		Active($pp_lkg_sheet,$PP_Book);
		$pp_lkg_sheet->Range("A1:AT57")->PasteSpecial(xlPasteAll);
		$Lkg_Book->Close();
		
		
		Active($pp_sheet,$PP_Book);
		#######################################Define the Headers###############
		my $Range; 
		Merger($pp_sheet,"A1:F1","Values from EMON");
		Merger($pp_sheet,"G1:H1","Temperature");
		Merger($pp_sheet,"I1:P1","IA C-State Residency");
		Merger($pp_sheet,"Q1:S1","DRAM Bandwidth");
		Merger($pp_sheet,"T1:Y1","Post Processed EMON Data");
		$pp_sheet->Range("T1:Y1")->{Interior}->{Color}=5296274;	
		$Range = "Z1:AB1";
		Merger($pp_sheet,$Range,"Core 0 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;	
		$Range = "AC1:AE1";Merger($pp_sheet,$Range,"Core 1 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;	
		$Range = "AF1:AH1";Merger($pp_sheet,$Range,"Core 2 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;	
		$Range = "AI1:AK1";Merger($pp_sheet,$Range,"Core 3 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;

		
		
		$Range = "AL1:AO1";Merger($pp_sheet,$Range,"DRAM Bandwidth");$pp_sheet->Range("$Range")->{Interior}->{Color}=192;
		$Range = "AP1:AU1";Merger($pp_sheet,$Range,"Measured Power");$pp_sheet->Range("$Range")->{Interior}->{Color}=49407;		

		Merger($pp_sheet,"AV1:BJ1","5 Sec Average");	
		Merger($pp_sheet,"BK1:BM1","5 Sec Average");
		$pp_sheet->Range("BN1")->{Value} = "Cdyn MAX";
		$pp_sheet->Range("BN1")->{Borders}->{Weight}=xlMedium;
		Merger($pp_sheet,"BO1:BP1","App Ratio");
		
		$Range = "BQ1:BT1";Merger($pp_sheet,$Range,"IA C0 Residency - 5 sec AVG");$pp_sheet->Range("$Range")->{Interior}->{Color}=12611584;		
		$Range = "BV1:BY1";Merger($pp_sheet,$Range,"IA C6/C7 Residency - 5 sec AVG");$pp_sheet->Range("$Range")->{Interior}->{Color}=12611584;		
		$Range = "CA1:CD1";Merger($pp_sheet,$Range,"IA C3 Residency - 5 sec AVG");$pp_sheet->Range("$Range")->{Interior}->{Color}=12611584;		

		Merger($pp_sheet,"CF1:CI1","DRAM Bandwidth (GB/s)");	
		Merger($pp_sheet,"CJ1:CL1","LLC Bandwidth (GB/s)");	
		
		$Range = "CM1:CT1";Merger($pp_sheet,$Range,"Package Power Rails");$pp_sheet->Range("$Range")->{Interior}->{Color}=49407;		
		$Range = "CU1:CX1";Merger($pp_sheet,$Range,"5 Second Average - Package Power Rails");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;		
		
		##################### Second Row Updated for Headers##############
		Row_2($pp_sheet);	
		
		my $gt_freq = 1000000000;
		#######################Formulas Loading##########################
		my $temp = $e_last[0]<=$d_last[0]?$e_last[0]:$d_last[0];
		$pp_sheet->Range("A3:D$temp")->{FormulaR1C1}="=EMON!R[-1]C";  ##EMON Values
		$pp_sheet->Range("E3:H$temp")->{FormulaR1C1}="=IF(EMON!R[-1]C=0,R[-1]C,EMON!R[-1]C)"; ##EMON Temperature and Frequency
		$pp_sheet->Range("T3")->{Value}=1;
		$pp_sheet->Range("T4:T$temp")->{FormulaR1C1}="=R[-1]C+1"; 
		$pp_sheet->Range("U3:W$temp")->{FormulaR1C1}="=RC[-19]/1000"; ##Reported Power From EMON
		$pp_sheet->Range("X3:X$temp")->{FormulaR1C1}="=RC[-19]/2*100"; ##Reported GT Frequency From EMON
		$pp_sheet->Range("Y3:Y$temp")->{FormulaR1C1}="=RC[-19]*100";  ##Reported IA Frequency From EMON
		$pp_sheet->Range("AP3")->{Value}=1;		
		$pp_sheet->Range("AP4:AP$temp")->{FormulaR1C1}="=R[-1]C+1";
		$pp_sheet->Range("AQ3:AR$temp")->{FormulaR1C1}="=NI_POWER!R[-1]C[-40]"; ## Importing the NI Data into the Main Sheet
		$pp_sheet->Range("AV3:AV$temp")->{FormulaR1C1}="=AVERAGE(RC[-4]:R[+49]C[-4])";
		$pp_sheet->Range("AW3:AW$temp")->{FormulaR1C1}="=AVERAGE(RC[-6]:R[+49]C[-6])";
		$pp_sheet->Range("AX3:AX$temp")->{FormulaR1C1}="=AVERAGE(RC[-29]:R[+49]C[-29])";
		$pp_sheet->Range("AY3:AY$temp")->{FormulaR1C1}="=ROUND(AVERAGE(RC[-43]:R[+49]C[-43]),0)";#my $a1 = 11;my $a2 = 56;
		$pp_sheet->Range("AZ3:AZ$temp")->{FormulaR1C1}="=VLOOKUP(RC[-1],Leakage_Lookup!R11C21:R56C26,6)*(RC[-4]/VLOOKUP(RC[-1],Leakage_Lookup!R11C21:R56C29,9))^3.7";
		$pp_sheet->Range("BK3:BK$temp")->{FormulaR1C1}="=(RC[-14]-RC[-11])/(RC[-15]*RC[-15]*$gt_freq)*10^9";
		$pp_sheet->Range("DF3")->{Formula}="=PERCENTILE.EXC(AQ3:AQ$temp,0.7)";
		$pp_sheet->Range("DG3:DG$temp")->{FormulaR1C1}="=(R3C110-RC[-68])/R3C110*100";
		$pp_sheet->Range("DH3:DH$temp")->{FormulaR1C1}="=IF(RC[-1]<20,1,0)";
		my $stop = $pp_sheet->Range("DH3:DH$temp")->Find({What=>"1",LookIn=>xlValues,SearchDirection=>xlPrevious,SearchOrder=>xlByRows})->{Row};
		my $start = $pp_sheet->Range("DH3:DH$temp")->Find({What=>"1",LookIn=>xlValues,SearchDirection=>xlNext,SearchOrder=>xlByRows})->{Row};
        #print "Start = $start  Stop = $stop \n";	<STDIN>;
		# $pp_sheet->Range("AR3:AR$temp")->{FormulaR1C1}="=NI_POWER!RC[-36]";
		
		#my $Series1 = $pp_sheet->Range("BK3:BK$temp");
		my $Series1 = $pp_sheet->Range("AQ3:AQ$temp");
		
		my $Chart = $Excel->Charts->Add;
		$Chart->{ChartType}=xlLine;
		$Chart->{Name}="CDyn Plot";
		$Chart->SetSourceData({Source=>$Series1,PlotBy=>xlColumns});
		$Chart->SeriesCollection(1)->Select;my $s1 =  "Cdyn";my $s2 =  "GT_Power";
		$Chart->SeriesCollection(1)->{Name}="=Data!U2";
		#$Chart->SeriesCollection(1)->Points($start)->Format->Fill->ForeColor->{RGB}="=RGB(0,255,0)";
		#$Chart->SeriesCollection(1)->Points($stop)->Format->Fill->ForeColor->{RGB}="=RGB(0,255,0)";
		$Chart->SeriesCollection->NewSeries;
		$Chart->SeriesCollection(2)->Select;
		$Chart->SeriesCollection(2)->{Name}="=Data!BK2";
		$Chart->SeriesCollection(2)->{Values}="=Data!BK3:BK$temp";
		
		$Chart->{HasTitle} = 1;
		$Chart->ChartTitle->{Text}="CDyn Plot";
		$pp_ave_sheet->Range("C7")->{VALUE}="Average Cdyn";$pp_ave_sheet->Range("C7:D8")->{INTERIOR}->{COLOR}=65535;$pp_ave_sheet->Range("C7:D8")->{Borders}->{Weight}=xlThin;
		$pp_ave_sheet->Range("C8")->{VALUE}="Average GT_Power"; #$pp_ave_sheet->Range("C8")->{INTERIOR}->{COLOR}=65535;$pp_sheet->Range("1")->{Borders}->{Weight}=xlThin;
		$pp_ave_sheet->Range("D7")->{FORMULA}="=AVERAGE(Data!BK$start:BK$stop";
		$pp_ave_sheet->Range("D8")->{FORMULA}="=AVERAGE(Data!AW$start:AW$stop";
		
		$pp_ave_sheet->Range("J1")->{VALUE}="CPU Frequency";$pp_ave_sheet->Range("J1:J3")->{INTERIOR}->{COLOR}=65535;$pp_ave_sheet->Range("J1:K3")->{Borders}->{Weight}=xlThin;
        $pp_ave_sheet->Range("K1")->{VALUE}="3200000000";
        $pp_ave_sheet->Range("J2")->{VALUE}="GT_FrequencyFrequency";
        $pp_ave_sheet->Range("K2")->{VALUE}="$gt_freq";
        $pp_ave_sheet->Range("J3")->{VALUE}="GT_Cdyn_VIRUS";
        $pp_ave_sheet->Range("K3")->{VALUE}="GT_Cdyn_VIRUS";
        $pp_ave_sheet->Range("N1")->{VALUE}="START";
        $pp_ave_sheet->Range("O1")->{VALUE}="$start";
        $pp_ave_sheet->Range("N2")->{VALUE}="END";
        $pp_ave_sheet->Range("O2")->{VALUE}="$stop";
        
        my $a_cdyn = $pp_ave_sheet->Range("D7")->{VALUE};
		my $gt_pwr = $pp_ave_sheet->Range("D8")->{VALUE};
        print "$name,$a_cdyn,$gt_pwr,$start,$stop\n";
        my $out_sheet = "c:\\Users\\ngupta5\\Cdyn_Part2\\".$name."\.xlsx";
        $PP_Book=$temp_str;
		$PP_Book -> SaveAs({Filename=>"$out_sheet"});#, FileFormat=>xlOpenXMLWorkbook});
#		$PP_Book->Close();
       


	}
}	
my $dir = getcwd();
if($dir=~s/\//\\\\/g)
{}
find(\&wanted, $dir);

=begin GHOSTCODE      
=end GHOSTCODE
=cut
  
    
