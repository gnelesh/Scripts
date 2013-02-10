use strict;
use File::Find;
use Cwd;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';
my @ARR;my $i=1;$ARR[0]='';
my $Book1;    my $raw="Raw";
my @A=(0..100);
my @ARR;
my @row2 = ("Time","GT_Power (0)","IA_Power (0)","PKG_Power (0)","GT_Freq (0)","IA_Freq (0)","IA_temp (0)","GT_temp (0)","IA_C7_res (0)","IA_C7_res (1)","IA_C7_res (2)","IA_C7_res (3)","IA_C7_res (4)","IA_C7_res (5)","IA_C7_res (6)","IA_C7_res (7)","IA_C6_res (0)","IA_C6_res (1)","IA_C6_res (2)","IA_C6_res (3)","IA_C6_res (4)","IA_C6_res (5)","IA_C6_res (6)","IA_C6_res (7)","IA_C3_res (0)","IA_C3_res (1)","IA_C3_res (2)","IA_C3_res (3)","IA_C3_res (4)","IA_C3_res (5)","IA_C3_res (6)","IA_C3_res (7)","RC6","RC6p","GT_DRAM_BW (0)","IA_DRAM_BW (0)","IO_DRAM_BW (0)","GT_UC_BW (0)","IA_UC_BW (0)","IA_UC_BW (1)","IA_UC_BW (2)","IA_UC_BW (3)","LLC_HITS0 (0)","LLC_HITS1 (0)","LLC_HITS2 (0)","LLC_HITS3 (0)","LLC_MISS0 (0)","LLC_MISS1 (0)","LLC_MISS2 (0)","LLC_MISS3 (0)","Time","GT Power","IA Power","PKG Power","GT Freq","IA Freq","C6 - C7","C3","C0","C6 - C7","C3","C0","C6 - C7","C3","C0","C6 - C7","C3","C0","GT_DRAM_BW (0)","IA_DRAM_BW (0)","IO_DRAM_BW (0)","Total DRAM BW","Time (seconds)","Gfx Power","GFX Vcc","IA Power","IA Vcc","Package Power","Gfx Vcc","Measured GT Power","Reported GT Power","GT Temperature","GT Leakage","IA Vcc","Measured IA Power","Reported IA Power","IA Temperature","IA + Ring/LLC Leakage","Ring/LLC Leakage","IA C6/C7 Residency","IA core Only Leakage","Measured Package Power","Reported Package Power","GT Cdyn","IA + Ring Cdyn","GT+IA Cdyn","GT Max Cdyn","GT App Ratio","IA App Ratio","IA Core0","IA Core1","IA Core2","IA Core3","IA 2 core C0 Avg","IA Core0","IA Core1","IA Core2","IA Core3","IA 2 core C6 Avg","IA Core0","IA Core1","IA Core2","IA Core3","IA 2 core C3 Avg","RC6","RC6p  - RC6plus - Power Gated","RC0","GT","IA","IO Bandwidth","Total","Uncacheable BW","LLC_HITS0 (0)","LLC_HITS1 (0)","LLC_HITS2 (0)","LLC_HITS3 (0)","LLC_MISS0 (0)","LLC_MISS1 (0)","LLC_MISS2 (0)","LLC_MISS3 (0)","IA LLC BW","GT LLC BW","Total LLC BW","SA Power","SA Vcc","VDDQ Power","VDDQ Vcc","VCCP (Uncore) Power","VCCP (Uncore) Vcc","SFR Power","SFR Vcc","SA ","VDDQ ","VCCP ","SFR","IA Total C0 ( Sum of All core C0 residency)","IA Total C3 ( Sum of All core C3 residency)","IA Total C6 ( Sum of All core C6 residency)");
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

sub Merger
{
	#print "$_[1]\n";
	$_[0]->Range("$_[1]")->Merge;
	$_[0]->Range("$_[1]")->{Value}="$_[2]";
	$_[0]->Range("$_[1]")->{HorizontalAlignment}= xlHAlignCenter;
	$_[0]->Range("$_[1]")->{Borders}->{Weight}=xlMedium;
}

sub Active
{
	my $na = $_[0]->{Name};
    #print "Activating --> $na\n";
	$_[1] =$_[0]->Activate();
	#$_[0]->Worksheets($name)->Activate();
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
			my $Emon=$emon_arr[0];
			my $path =getcwd();
			if ($path=~s/\//\\\\/g)
			{}
			my $abs_path_e = $path.'\\'.$Emon;
			my $name;
			if ($Emon =~/(\S+)\.csv/)
			{
				$name =  $1;
			}
		my $Emon_Book = $Excel->Workbooks->Open("$abs_path_e");
		my $Emon_Sheet = $Emon_Book->Worksheets(1);
		my @e_last = Last($Emon_Sheet);
		$Excel->{SheetsInNewWorkBook}=1;
		my $PP_Book=$Excel->Workbooks->Add();my $temp_str = $PP_Book;
									
		####################################### AVERAGE SHEET ###########################
		
		$PP_Book->Worksheets->Add;
		my $pp_ave_sheet=$PP_Book->Worksheets(1);
		$pp_ave_sheet->{Name}="Experiment_Summary";
		
	   ######################################## EMON SHEET #############################
		
		$PP_Book->Worksheets->Add;
		my $pp_emon_sheet=$PP_Book->Worksheets(1);
		$pp_emon_sheet->{Name}="EMON";
		
		####################################### Post Process SHEET #####################
		$PP_Book->Worksheets->Add;
		my $pp_sheet=$PP_Book->Worksheets(1);
		$pp_sheet->{Name}="Data";
		
		foreach my $k(1..3)
		{
			my $name= $PP_Book->Worksheets($k)->{Name};
            print "$k-->$name\n";
		}
		#Active($Emon_Book,$Emon_Sheet);
		$Emon_Sheet->Range("A1:$ARR[$e_last[1]]$e_last[0]")->Copy();
		Active($pp_emon_sheet,$PP_Book);
		$pp_emon_sheet->Range("A1:$ARR[$e_last[1]]$e_last[0]")->PasteSpecial();
		$Emon_Book->Close();
		
		#######################################Define the Headers###############
		my $Range; 
		Merger($pp_sheet,"A1:F1","Values from EMON");
		Merger($pp_sheet,"G1:H1","Temperature");
		Merger($pp_sheet,"I1:P1","IA C7-State Residency");
		Merger($pp_sheet,"Q1:X1","IA C6-State Residency");
		Merger($pp_sheet,"Y1:AF1","IA C3-State Residency");
		Merger($pp_sheet,"AG1:AH1","GT RC-State Residency");
		Merger($pp_sheet,"AI1:AK1","DRAM Bandwidth");
		Merger($pp_sheet,"AL1:AW1","LLC BW Measurements");
		Merger($pp_sheet,"AY1:BD1","Post Processed EMON Data");
		$pp_sheet->Range("AY1:BD1")->{Interior}->{Color}=5296274;	
		$Range = "BE1:BG1";Merger($pp_sheet,$Range,"Core 0 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;	
		$Range = "BH1:BJ1";Merger($pp_sheet,$Range,"Core 1 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;	
		$Range = "BK1:BM1";Merger($pp_sheet,$Range,"Core 2 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;	
		$Range = "BN1:BP1";Merger($pp_sheet,$Range,"Core 3 IA C-State Residency");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;

		
		
		$Range = "BQ1:BT1";Merger($pp_sheet,$Range,"DRAM Bandwidth");$pp_sheet->Range("$Range")->{Interior}->{Color}=192;
		$Range = "BU1:BZ1";Merger($pp_sheet,$Range,"Measured Power");$pp_sheet->Range("$Range")->{Interior}->{Color}=49407;		

		Merger($pp_sheet,"CA1:CO1","5 Sec Average");	
		Merger($pp_sheet,"CP1:CR1","CDyn");
		$pp_sheet->Range("CS1")->{Value} = "Cdyn MAX";
		$pp_sheet->Range("CS1")->{Borders}->{Weight}=xlMedium;
		Merger($pp_sheet,"CT1:CU1","App Ratio");
		
		$Range = "CV1:CY1";Merger($pp_sheet,$Range,"IA C0 Residency - 5 sec AVG");$pp_sheet->Range("$Range")->{Interior}->{Color}=12611584;		
		$Range = "DA1:DD1";Merger($pp_sheet,$Range,"IA C6/C7 Residency - 5 sec AVG");$pp_sheet->Range("$Range")->{Interior}->{Color}=12611584;		
		$Range = "DF1:DI1";Merger($pp_sheet,$Range,"IA C3 Residency - 5 sec AVG");$pp_sheet->Range("$Range")->{Interior}->{Color}=12611584;		

		Merger($pp_sheet,"DK1:DM1","GT C-State Residency");	
		Merger($pp_sheet,"DN1:DQ1","DRAM Bandwidth (GB/s)");	
		Merger($pp_sheet,"DR1:EC1","LLC Bandwidth (GB/s)");	
		
		$Range = "ED1:EK1";Merger($pp_sheet,$Range,"Package Power Rails");$pp_sheet->Range("$Range")->{Interior}->{Color}=49407;		
		$Range = "EL1:EO1";Merger($pp_sheet,$Range,"5 Second Average - Package Power Rails");$pp_sheet->Range("$Range")->{Interior}->{Color}=15773696;		
		
		##################### Second Row Updated for Headers##############
		Row_2($pp_sheet);	
		
		my $gt_freq = 1000000000;
		my $temp = $e_last[0];
		$pp_sheet->Range("A3:AH$temp")->{FormulaR1C1}="=EMON!R[-1]C";  ##EMON Values
		$pp_sheet->Range("AI3:AK$temp")->{FormulaR1C1}="=EMON!R[-1]C[+1]"; ##DRAM Bandwidth
		$pp_sheet->Range("AY3")->{Value}=1;
		$pp_sheet->Range("AY4:AY$temp")->{FormulaR1C1}="=R[-1]C+1";
		$pp_sheet->Range("AZ3:BB$temp")->{FormulaR1C1}="=RC[-50]/1000"; ##Reported Power From EMON
		$pp_sheet->Range("BC3:BC$temp")->{FormulaR1C1}="=IF((RC[-50]/2*100)=0,R[-1]C,RC[-50]/2*100)"; ##Reported GT Frequency From EMON
		$pp_sheet->Range("BD3:BD$temp")->{FormulaR1C1}="=IF((RC[-50]*100)=0,R[-1]C,RC[-50]*100)";  ##Reported IA Frequency From EMON
		
		#################################################################Core 0 IA C State Residency ########################################################################
		$pp_sheet->Range("BE3:BE$temp")->{FormulaR1C1}="=(IF(((R[+1]C[-40]-RC[-40])/0.1/(RC[-1]*10^6))>1,1,(R[+1]C[-40]-RC[-40])/0.1/(RC[-1]*10^6))+IF(((R[+1]C[-39]-RC[-39])/0.1/(RC[-1]*10^6))>1,1,(R[+1]C[-39]-RC[-39])/0.1/(RC[-1]*10^6)))/2";
		$pp_sheet->Range("BF3:BF$temp")->{FormulaR1C1}="=(((R[+1]C[-33]-RC[-33])/0.1/(RC[-2]*10^6))+((R[+1]C[-32]-RC[-32])/0.1/(RC[-2]*10^6)))/2";
		$pp_sheet->Range("BG3:BG$temp")->{FormulaR1C1}="=IF(1-SUM(RC[-2]:RC[-1])>0,1-SUM(RC[-2]:RC[-1]),0)";
		
		################################################################# Core 1 IA C State Residency ########################################################################
		$pp_sheet->Range("BH3:BH$temp")->{FormulaR1C1}="=(IF(((R[+1]C[-41]-RC[-41])/0.1/(RC[-4]*10^6))>1,1,(R[+1]C[-41]-RC[-41])/0.1/(RC[-4]*10^6))+IF(((R[+1]C[-40]-RC[-40])/0.1/(RC[-4]*10^6))>1,1,(R[+1]C[-40]-RC[-40])/0.1/(RC[-4]*10^6)))/2";
		$pp_sheet->Range("BI3:BJ$temp")->{FormulaR1C1}="=(((R[+1]C[-34]-RC[-34])/0.1/(RC[-5]*10^6))+((R[+1]C[-33]-RC[-33])/0.1/(RC[-5]*10^6)))/2";
		$pp_sheet->Range("BJ3:BJ$temp")->{FormulaR1C1}="=IF(1-SUM(RC[-2]:RC[-1])>0,1-SUM(RC[-2]:RC[-1]),0)";
		
		################################################################# Core 2 IA C State Residency ########################################################################
		$pp_sheet->Range("BK3:BK$temp")->{FormulaR1C1}="=(IF(((R[+1]C[-42]-RC[-42])/0.1/(RC[-7]*10^6))>1,1,(R[+1]C[-42]-RC[-42])/0.1/(RC[-7]*10^6))+IF(((R[+1]C[-41]-RC[-41])/0.1/(RC[-7]*10^6))>1,1,(R[+1]C[-41]-RC[-41])/0.1/(RC[-7]*10^6)))/2";
		$pp_sheet->Range("BL3:BL$temp")->{FormulaR1C1}="=(((R[+1]C[-35]-RC[-35])/0.1/(RC[-8]*10^6))+((R[+1]C[-34]-RC[-34])/0.1/(RC[-8]*10^6)))/2";
		$pp_sheet->Range("BM3:BM$temp")->{FormulaR1C1}="=IF(1-SUM(RC[-2]:RC[-1])>0,1-SUM(RC[-2]:RC[-1]),0)";
		
		################################################################# Core 3 IA C State Residency ########################################################################
		$pp_sheet->Range("BN3:BN$temp")->{FormulaR1C1}="=(IF(((R[+1]C[-43]-RC[-43])/0.1/(RC[-10]*10^6))>1,1,(R[+1]C[-43]-RC[-43])/0.1/(RC[-10]*10^6))+IF(((R[+1]C[-42]-RC[-42])/0.1/(RC[-10]*10^6))>1,1,(R[+1]C[-42]-RC[-42])/0.1/(RC[-10]*10^6)))/2";
		$pp_sheet->Range("BO3:BO$temp")->{FormulaR1C1}="=(((R[+1]C[-36]-RC[-36])/0.1/(RC[-11]*10^6))+((R[+1]C[-35]-RC[-35])/0.1/(RC[-11]*10^6)))/2";
		$pp_sheet->Range("BP3:BP$temp")->{FormulaR1C1}="=IF(1-SUM(RC[-2]:RC[-1])>0,1-SUM(RC[-2]:RC[-1]),0)";
		
		################################################################# DRAM Bandwidth ########################################################################
		
		$pp_sheet->Range("BQ3:BS$temp")->{FormulaR1C1}="=IF((R[+1]C[-34]>=RC[-34]),(((R[+1]C[-34]-RC[-34])*64/0.1)/1024^3),(((4294967295-R[+1]C[-34]+RC[-34])*64/0.1)/1024^3))";
		$pp_sheet->Range("BT3:BT$temp")->{FormulaR1C1}="=SUM(RC[-3]:RC[-1])";
		
		
		################################################################# 5 Sec Averages ########################################################################
				
				
				####################################################			C0 		 ########################################
		$pp_sheet->Range("CV3:CV$temp")->{FormulaR1C1}="=AVERAGE(RC[-41]:R[+49]C[-41])";
		$pp_sheet->Range("CW3:CW$temp")->{FormulaR1C1}="=AVERAGE(RC[-39]:R[+49]C[-39])";
		$pp_sheet->Range("CX3:CX$temp")->{FormulaR1C1}="=AVERAGE(RC[-37]:R[+49]C[-37])";
		$pp_sheet->Range("CY3:CY$temp")->{FormulaR1C1}="=AVERAGE(RC[-35]:R[+49]C[-35])";
		$pp_sheet->Range("CZ3:CZ$temp")->{FormulaR1C1}="=AVERAGE(RC[-4]:RC[-1])";
		
				####################################################      	C6/7     #######################################
		$pp_sheet->Range("DA3:DA$temp")->{FormulaR1C1}="=AVERAGE(RC[-48]:R[+49]C[-48])";
		$pp_sheet->Range("DB3:DB$temp")->{FormulaR1C1}="=AVERAGE(RC[-46]:R[+49]C[-46])";
		$pp_sheet->Range("DC3:DC$temp")->{FormulaR1C1}="=AVERAGE(RC[-44]:R[+49]C[-44])";
		$pp_sheet->Range("DD3:DD$temp")->{FormulaR1C1}="=AVERAGE(RC[-42]:R[+49]C[-42])";
		$pp_sheet->Range("DE3:DE$temp")->{FormulaR1C1}="=AVERAGE(RC[-4]:RC[-1])";
				####################################################      	C3     		#######################################
		$pp_sheet->Range("DF3:DF$temp")->{FormulaR1C1}="=AVERAGE(RC[-52]:R[+49]C[-52])";
		$pp_sheet->Range("DG3:DG$temp")->{FormulaR1C1}="=AVERAGE(RC[-50]:R[+49]C[-50])";
		$pp_sheet->Range("DH3:DH$temp")->{FormulaR1C1}="=AVERAGE(RC[-48]:R[+49]C[-48])";
		$pp_sheet->Range("DI3:DI$temp")->{FormulaR1C1}="=AVERAGE(RC[-46]:R[+49]C[-46])";
		$pp_sheet->Range("DJ3:DJ$temp")->{FormulaR1C1}="=AVERAGE(RC[-4]:RC[-1])";
		
		####################################################      			GT RC STATE RESIDENCY     		#######################################
		
		$pp_sheet->Range("DK3:DK$temp")->{FormulaR1C1}="=IF(((R[+1]C[-82]-RC[-82])*(1.28*10^-6)/0.1)>1,1,((R[+1]C[-82]-RC[-82])*(1.28*10^-6)/0.1))";
		$pp_sheet->Range("DL3:DL$temp")->{FormulaR1C1}="=IF(((R[+1]C[-81]-RC[-81])*(1.28*10^-6)/0.1)>1,1,((R[+1]C[-81]-RC[-81])*(1.28*10^-6)/0.1))";
		$pp_sheet->Range("DM3:DM$temp")->{FormulaR1C1}="=1-RC[-1]";
		
		
		###################################################################   DRAM ############################################################
		
		$pp_sheet->Range("DN3:DQ$temp")->{FormulaR1C1}="=AVERAGE(RC[-49]:R[+49]C[-49])";
		
		
		
		
		$pp_sheet->Range("BE3:BP$temp")->{NumberFormat}="0.00%";
		$pp_sheet->Range("CV3:DM$temp")->{NumberFormat}="0.00%";
		
		}
	}
	my $dir = getcwd();
	if ($dir=~s/\//\\\\/g)
	{}
	find(\&wanted,$dir);
		
		
	