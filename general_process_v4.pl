#!/usr/bin/perl
#Contact: caoshuhuan@catb.org.cn
#date:Fri Jan 26 09:44:17 CST 2018
use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseXLSX;
use Encode;
use utf8;
use Getopt::Std;

my $version = '0.4.26';
my %opts;
getopts("S:d:I:c:t:o:h",\%opts);

my $sampleinfo = $opts{S};	#样本信息.xlsx
my $datadpth = $opts{d};	#各样本数据量理论深度.xlsx
my $Interdir = $opts{I};	#样本结果文件目录
#my $InterSum = $opts{A};	#Interface summary of all samples
my $caseInterdir = $opts{c};	#case-control||case-case||pedigree||case-case-control Interface directory which contains case-control||case-case Interface_summary.txt
my $type = $opts{t};		#不同的分析模式，比如单样本，case-case等
my $outdir = $opts{o};		#保存根据不同的样本编号生成不同的格式化结果的目录
#my $totalfile = $opts{s};	#样本突变情况统计
#my $meandir = $opts{m};	#样本深度覆盖度情况结果目录


&print_usage unless (defined($sampleinfo));
&print_usage unless (defined($datadpth));
#&print_usage unless (defined($InterSum));
&print_usage unless (defined($outdir));
&print_usage unless (defined($Interdir));
#&print_usage unless (defined($totalfile));
#&print_usage unless (defined($meandir));
&print_usage if (defined($opts{h}));

my $typenum = 1 if ($type eq 'S');	#single
$typenum    = 2 if ($type eq 'C');	#case-control
$typenum    = 3 if ($type eq 'CC');	#case-case
$typenum    = 4 if ($type eq 'P');	#pedigree
$typenum    = 5 if ($type eq 'CCC');#case-case-control目前设定3个，2个case1个control

if ($typenum == 1){
	my $parser = Spreadsheet::ParseXLSX->new();
	$sampleinfo=decode("GB2312",$sampleinfo);
	$sampleinfo=encode("GB2312",$sampleinfo);
	#print $sampleinfo;
	my $workbook = $parser ->parse("$sampleinfo");
	
	my($receive_date,$sampleID,$patientname,$donorname,$kiname,$age,$gender,$sampletype,$diagnose,$hospital,$sampleQ,$sequenceQ,$Q30);
	my(%receive_date_hash,%patientname_hash,%donorname_hash,%age_hash,%gender_hash,%sampletype_hash,%diagnose_hash,%hospital_hash,%sampleQ_hash,%sequenceQ_hash,%Q30_hash);
	my(@sampleIDs,$InterSum);
	
	for my $worksheet($workbook->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$receive_date = $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+1);
			$sampleID=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$patientname=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+3);
			$donorname=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+4);
			$kiname=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+5);
			$age=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+6);
			$gender=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+7);
			$sampletype=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+9);
			$diagnose=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+11);
			$hospital=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+14);
			$sampleQ=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+15);
			$sequenceQ=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+16);
			$Q30=$cell? encode ('GB2312', $cell->value()) :'';
			if($sampleID){
				$receive_date_hash{$sampleID}= $receive_date;
				$patientname_hash{$sampleID} = $patientname;
				$donorname_hash{$sampleID}   = $donorname;
				$age_hash{$sampleID}         = $age;
				$gender_hash{$sampleID}      = $gender;
				$sampletype_hash{$sampleID}  = $sampletype;
				$diagnose_hash{$sampleID}    = $diagnose;
				$hospital_hash{$sampleID}    = $hospital;
				$sampleQ_hash{$sampleID}     = $sampleQ;
				$sequenceQ_hash{$sampleID}   = $sequenceQ;
				$Q30_hash{$sampleID}         = $Q30;
			}
		#print $receive_date,"\t",$sampleID,"\t",$patientname,"\n";
		push @sampleIDs,$sampleID;
		}
	}
	
	my $parser = Spreadsheet::ParseXLSX->new();
	$datadpth=decode("GB2312",$datadpth);
	$datadpth=encode("GB2312",$datadpth);
	
	my $workbook = $parser ->parse("$datadpth");
	my ($sample_id,$datasize,$theodepth);
	my (%q30_hash,%datasize_hash,%theodepth_hash);
	
	for my $worksheet($workbook->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$sample_id = $cell? encode ('GB2312', $cell->value()) :'';
			#my $cell = $worksheet->get_cell($row,$col_min+1);
			#$q30=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+1);
			$datasize=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$theodepth=$cell? encode ('GB2312', $cell->value()) :'';
			if($sample_id){
				#$q30_hash{$sample_id} = $q30;
				$datasize_hash{$sample_id} = $datasize;
				$theodepth_hash{$sample_id}= $theodepth;
			}
		}
	}
	
	if (defined ($opts{o})){
		$outdir .='\\' unless ($outdir=~/\\$/);
	}
	if(defined ($opts{I})){
		$Interdir .='\\' unless ($Interdir=~/\\$/);
	}
	if(defined ($opts{c})){
		$caseInterdir .='\\' unless ($caseInterdir=~/\\$/);
	}
	my(%t4_hash,%t6_hash,%varNum_hash,%MeanDp_hash,%twentyCvg_hash,%tenDp_hash,%fiftyDp_hash,%hundredDp_hash);
	
	opendir my $dh,$Interdir;
	my @interfiles = readdir($dh);
	my(%vartypefile_hash,%dpcvgfile_hash,%pharm_hash,%T4file_hash,%T6file_hash);
	foreach my $file(@interfiles){
		if($file=~ /(\w+)\_(\w+)\_variant_types_number/){
			my $sampleid = $1;
			$vartypefile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_MeanDepth_Coverage/){
			my $sampleid = $1;
			$dpcvgfile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_Recalibrated.output.vcf/){
		#elsif($file=~ /(\w+)\_(\w+)\_DrugResponse.output.vcf/){	#新版本的文件名
			my $sampleid = $1;
			$pharm_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T4.txt/){
			my $sampleid = $1;
			$T4file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T6.txt/){
			my $sampleid = $1;
			$T6file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /Interface_summary.txt/){
			$InterSum = $file;
		}
		else{next;}
	}
	open my $fh,$InterSum;
	while(<$fh>){
		next if /^#/;
		chomp;
		my @tmp = split /\t/,$_;
		my $sampleid = $tmp[0];
		$varNum_hash{$sampleid} = $tmp[1];
		$t4_hash{$sampleid} = $tmp[2];
		$t6_hash{$sampleid} = $tmp[3];
		$MeanDp_hash{$sampleid} = $tmp[4];
		$twentyCvg_hash{$sampleid} = $tmp[5];
		$tenDp_hash{$sampleid} = $tmp[6];
		$fiftyDp_hash{$sampleid} = $tmp[7];
		$hundredDp_hash{$sampleid} = $tmp[8];
	}
	close $fh;
	
	foreach my $sampleid(@sampleIDs){
		my $destindir = $outdir.$sampleid.'_screen';
		system("mkdir $destindir") unless(-e $destindir);
		open INFO,'>',"$destindir\\"."sample_info.csv";
		open GEN,'>',"$destindir\\"."general.csv";
		open Q1,'>',"$destindir\\"."quality_1.csv";
		open T1,'>',"$destindir\\"."patients.csv";
		open T2,'>',"$destindir\\"."quality_2.csv";
		print Q1 encode("GB2312","样本编号\t样本质控\t测序质控\tQ30\n");
		print Q1 "$sampleid\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\n";
		print INFO "name\n";
		print INFO "sample.num\t$typenum\n";
		print INFO "sample.type\t$sampletype_hash{$sampleid}\n";
		print GEN "name\n";
		printf GEN "data\t%d\n",$datasize_hash{$sampleid};
		printf GEN "raw.depth\t%d\n",$theodepth_hash{$sampleid};
		print INFO "sample.code\t$sampleid\n";
		print GEN "total\t$varNum_hash{$sampleid}\n";
		print GEN "blood\t$t4_hash{$sampleid}\n";
		print GEN "blood.potential\t$t6_hash{$sampleid}\n";
		printf GEN "accuracy\t%.2f\n",$Q30_hash{$sampleid};
		print GEN "depth\t$MeanDp_hash{$sampleid}\n";						#理论深度
		#print INFO "seq.qual.3\t5%\n";										#突变比例
		print GEN "detect\t$twentyCvg_hash{$sampleid}\n";					#检出率
		print T1 encode("GB2312","样本编号\t$sampleid\n");
		print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\n";
		print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\n");
		print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\n";
		print T2 encode("GB2312","样本编号\t平均测序深度\t覆盖度10x\t覆盖度20x\t覆盖度50x\t覆盖度100x\n");
		print T2 $sampleid,"\t$MeanDp_hash{$sampleid}\t$tenDp_hash{$sampleid}\t$twentyCvg_hash{$sampleid}\t$fiftyDp_hash{$sampleid}\t$hundredDp_hash{$sampleid}\n";
		close INFO;
		close Q1;
		close T1;
		close T2;
		system("copy $vartypefile_hash{$sampleid} $destindir\\$sampleid\_table3.csv");
		#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
		my $t3file = "$destindir\\$sampleid\_table3.csv";
		open my $fh,$t3file;
		open OUT,'>',"$destindir\\mutants_statistics.csv";
		print OUT "20$sampleid\n";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			print OUT "$tmp[0]\t$tmp[1]\n";
		}
		close $fh;
		close OUT;
		system("del $destindir\\$sampleid\_table3.csv");
		system("copy $T4file_hash{$sampleid} $destindir\\$sampleid\_table4.csv");
		my $t4file = "$destindir\\$sampleid\_table4.csv";
		open my $fh,$t4file;
		open OUT,'>',"$destindir\\mutants_hot.csv";
		print OUT encode("GB2312","基因名\t变异描述\t突变比例\t致病性评估（仅供参考）\t其他信息\n");
		while(<$fh>){
			my @tmp = split /\t/,$_;
			$tmp[2] .= "%";
			@tmp = @tmp[0..$#tmp];
			my $new = join "\t",@tmp;
			print OUT $new;
		}
		close $fh;
		close OUT;
		system("del $destindir\\$sampleid\_table4.csv");
		system("copy $T6file_hash{$sampleid} $destindir\\$sampleid\_table6.csv");
		my $t6file = "$destindir\\$sampleid\_table6.csv";
		open my $fh,$t6file;
		open OUT,'>',"$destindir\\mutants_potential.csv";
		print OUT encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		my @unsorted_lines;
		while(<$fh>){	#超出20行的输出前20行，以num of harm按从小到大排，不足的全输
			chomp;
			push @unsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@unsorted_lines);
		for(@sorted_lines){
			print OUT encode("GB2312",$_);
		}
		close OUT;
		system("del $destindir\\$sampleid\_table6.csv");
		system("copy $pharm_hash{$sampleid} $destindir\\$sampleid\_table5.csv");
		my $t5file = "$destindir\\$sampleid\_table5.csv";
		open my $fh,$t5file;
		open OUT,'>',"$destindir\\mutants_medicine.csv";
		#print OUT encode("GB2312","化疗药物\t检测基因\t检测区域\t检测结果\t结果解读\t等级\n");
		#print OUT "化疗药物\t检测基因\t检测区域\t检测结果\t结果解读\t等级\n";
		print OUT "检测基因\t检测区域\t检测结果\t结果解读\t化疗药物\t等级\n";
		while(<$fh>){
			print OUT $_;
		}
		close $fh;
		close OUT;	
		system("del $destindir\\$sampleid\_table5.csv")
	}
}
elsif ($typenum == 2){							#case-control
	my $parser = Spreadsheet::ParseXLSX->new();
	$sampleinfo=decode("GB2312",$sampleinfo);
	$sampleinfo=encode("GB2312",$sampleinfo);
	#print $sampleinfo;
	my $workbook = $parser ->parse("$sampleinfo");
	
	my($receive_date,$sampleID,$patientname,$donorname,$kiname,$age,$gender,$sampletype,$diagnose,$hospital,$sampleQ,$sequenceQ,$Q30);
	my(%receive_date_hash,%patientname_hash,%donorname_hash,%age_hash,%gender_hash,%sampletype_hash,%diagnose_hash,%hospital_hash,%sampleQ_hash,%sequenceQ_hash,%Q30_hash);
	my(@sampleIDs,$InterSum);	#以后想办法以2行为单位读文件
	my(%ctrlID,%ctrltype,%ctrlQ,%ctrlseqQ,%ctrlQ30);
	for my $worksheet($workbook->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		my $cell = $worksheet->get_cell($row_min+1,$col_min);
		$receive_date = $cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+1);
		$sampleID=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+1);
		$ctrlID{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+2);
		$patientname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+3);
		$donorname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+4);
		$kiname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+5);
		$age=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+6);
		$gender=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+7);
		$sampletype=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+7);
		$ctrltype{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+9);
		$diagnose=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+11);
		$hospital=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+14);
		$sampleQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+14);
		$ctrlQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+15);
		$sequenceQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+15);
		$ctrlseqQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+16);
		$Q30=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+16);
		$ctrlQ30{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';
		if($sampleID){
			$receive_date_hash{$sampleID}= $receive_date;
			$patientname_hash{$sampleID} = $patientname;
			$donorname_hash{$sampleID}   = $donorname;
			$age_hash{$sampleID}         = $age;
			$gender_hash{$sampleID}      = $gender;
			$sampletype_hash{$sampleID}  = $sampletype;
			$diagnose_hash{$sampleID}    = $diagnose;
			$hospital_hash{$sampleID}    = $hospital;
			$sampleQ_hash{$sampleID}     = $sampleQ;
			$sequenceQ_hash{$sampleID}   = $sequenceQ;
			$Q30_hash{$sampleID}         = $Q30;
		}
		#print $receive_date,"\t",$sampleID,"\t",$patientname,"\n";
		#print "$sampleID\t$sampletype\t$ctrlID{$sampleID}\t$ctrltype{$sampleID}\n";
		push @sampleIDs,$sampleID;
	}
	
	my $parser = Spreadsheet::ParseXLSX->new();
	$datadpth=decode("GB2312",$datadpth);
	$datadpth=encode("GB2312",$datadpth);
	
	my $workbook = $parser ->parse("$datadpth");
	my (%datasize_hash,%theodepth_hash);
	
	for my $worksheet($workbook->worksheet(0)){
		my ($sample_id,$datasize,$theodepth);
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$sample_id = $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+1);
			$datasize=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$theodepth=$cell? encode ('GB2312', $cell->value()) :'';
			if($sample_id){
				#$q30_hash{$sample_id} = $q30;
				$datasize_hash{$sample_id} = $datasize;
				$theodepth_hash{$sample_id}= $theodepth;
			}
		}
	}
	#print "$sampleID\t$ctrlID{$sampleID}\t$datasize_hash{$sampleID}\t$datasize_hash{$ctrlID{$sampleID}}\n";
	if (defined ($opts{o})){
		$outdir .='\\' unless ($outdir=~/\\$/);
	}
	if(defined ($opts{I})){
		$Interdir .='\\' unless ($Interdir=~/\\$/);
	}
	if(defined ($opts{c})){
		$caseInterdir .='\\' unless ($caseInterdir=~/\\$/);
	}
	my(%t4_hash,%t6_hash,%varNum_hash,%MeanDp_hash,%twentyCvg_hash,%tenDp_hash,%fiftyDp_hash,%hundredDp_hash);
	
	opendir my $dh,$Interdir;
	my @interfiles = readdir($dh);
	my(%vartypefile_hash,%dpcvgfile_hash,%pharm_hash,%T4file_hash,%T6file_hash);
	foreach my $file(@interfiles){
		if($file=~ /(\w+)\_(\w+)\_variant_types_number/){
			my $sampleid = $1;
			$vartypefile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_MeanDepth_Coverage/){
			my $sampleid = $1;
			$dpcvgfile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_Recalibrated.output.vcf/){
		#elsif($file=~ /(\w+)\_(\w+)\_DrugResponse.output.vcf/){	#新版本的文件名
			my $sampleid = $1;
			$pharm_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T4.txt/){
			my $sampleid = $1;
			$T4file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T6.txt/){
			my $sampleid = $1;
			$T6file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /Interface_summary.txt/){
			$InterSum = $file;
		}
		else{next;}
	}
	
	open my $fh,$InterSum;
	while(<$fh>){
		next if /^#/;
		chomp;
		my @tmp = split /\t/,$_;
		my $sampleid = $tmp[0];
		$varNum_hash{$sampleid} = $tmp[1];
		$t4_hash{$sampleid} = $tmp[2];
		$t6_hash{$sampleid} = $tmp[3];
		$MeanDp_hash{$sampleid} = $tmp[4];
		$twentyCvg_hash{$sampleid} = $tmp[5];
		$tenDp_hash{$sampleid} = $tmp[6];
		$fiftyDp_hash{$sampleid} = $tmp[7];
		$hundredDp_hash{$sampleid} = $tmp[8];
	}
	close $fh;
	
	print STDERR "argument -c must be set after -T C\n" unless (defined $caseInterdir);
	opendir my $dh,$caseInterdir || die "can't open $caseInterdir\n";
	my(%spec_varNum_hash,%spec_t4_hash,%spec_t6_hash,%spec_T4file_hash,%spec_T6file_hash);
	my @caseinterfiles = readdir($dh);
	foreach my $file(@caseinterfiles){
		if ($file =~ /Interface_summary.txt/){
			open my $fh ,$file;
			while(<$fh>){
				next if /^#/;
				chomp;
				my @tmp = split /\t/,$_;
				my $sampleid = $tmp[0];
				$spec_varNum_hash{$sampleid} = $tmp[1];
				$spec_t4_hash{$sampleid} = $tmp[2];
				$spec_t6_hash{$sampleid} = $tmp[3];
			}
			close $fh;
		}
		if($file =~ /(\w+)\_(\w+)\_specific_T4.txt/){
			my $sampleid = $1;
			$spec_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+)\_specific_T6.txt/){
			my $sampleid = $1;
			$spec_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
	}
	closedir $dh;
	foreach my $sampleid(@sampleIDs){
		my $destindir = $outdir.$sampleid.'_'.$ctrlID{$sampleid}.'_case_ctrl';
		system("mkdir $destindir") unless(-e $destindir);
		open INFO,'>',"$destindir\\"."sample_info.csv";
		open GEN,'>',"$destindir\\"."general.csv";
		open Q1,'>',"$destindir\\"."quality_1.csv";
		open T1,'>',"$destindir\\"."patients.csv";
		open T2,'>',"$destindir\\"."quality_2.csv";
		print Q1 encode("GB2312","样本编号\t样本质控\t测序质控\tQ30\n");
		print Q1 "$sampleid\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\n";
		print Q1 "$ctrlID{$sampleid}\t$ctrlQ{$sampleID}\t$ctrlseqQ{$sampleID}\t$ctrlQ30{$sampleID}\n";
		print INFO "name\n";
		print INFO "sample.num\t$typenum\n";
		print INFO "sample.type.treat\t$sampletype_hash{$sampleid}\n";
		print INFO "sample.type.control\t$sampletype_hash{$ctrlID{$sampleid}}\n";
		print INFO "code.treat\t$sampleid\n";
		print INFO "code.control\t$ctrlID{$sampleid}\n";
		print GEN "name\n";
		my $totaldata = $datasize_hash{$sampleid}+$datasize_hash{$ctrlID{$sampleid}};
		printf GEN "data\t%d\n",$totaldata;
		printf GEN "raw.depth.treat\t%d\n",$theodepth_hash{$sampleid};
		printf GEN "raw.depth.control\t%d\n",$theodepth_hash{$ctrlID{$sampleid}};
		print GEN "total.treat.before\t$varNum_hash{$sampleid}\n";
		print GEN "total.control.before\t$varNum_hash{$ctrlID{$sampleid}}\n";
		print GEN "total.after\t$spec_varNum_hash{$sampleid}\n";								#对比后热点总个数
		print GEN "blood.before\t$t4_hash{$sampleid}\n";
		print GEN "blood.before.potential\t$t6_hash{$sampleid}\n";
		print GEN "blood.after\t$spec_t4_hash{$sampleid}\n";								#对比后血液病热点总个数
		print GEN "blood.after.potential\t$spec_t6_hash{$sampleid}\n";						#对比后潜在血液病热点总个数
		printf GEN "accuracy.treat\t%.2f\n",$Q30_hash{$sampleid};
		printf GEN "accuracy.control\t%.2f\n",$Q30_hash{$ctrlID{$sampleid}};
		print GEN "depth.treat\t$MeanDp_hash{$sampleid}\n";						#理论深度
		print GEN "depth.control\t$MeanDp_hash{$ctrlID{$sampleid}}\n";
		#print INFO "seq.qual.3\t5%\n";										#突变比例
		print GEN "detect.treat\t$twentyCvg_hash{$sampleid}\n";					#检出率
		print GEN "detect.control\t$twentyCvg_hash{$ctrlID{$sampleid}}\n";
		print T1 encode("GB2312","样本编号\t$sampleid\t$ctrlID{$sampleid}\n");
		#print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\t$patientname_hash{$ctrlID{$sampleid}}\n";
		print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\t$patientname_hash{$sampleid}\n";
		#print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\t$age_hash{$ctrlID{$sampleid}}\n");
		print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\t$age_hash{$sampleid}\n");
		#print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\t$gender_hash{$ctrlID{$sampleid}}\n";
		print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\t$gender_hash{$sampleid}\n";
		print T2 encode("GB2312","样本编号\t平均测序深度\t覆盖度10x\t覆盖度20x\t覆盖度50x\t覆盖度100x\n");
		print T2 $sampleid,"\t$MeanDp_hash{$sampleid}\t$tenDp_hash{$sampleid}\t$twentyCvg_hash{$sampleid}\t$fiftyDp_hash{$sampleid}\t$hundredDp_hash{$sampleid}\n";
		print T2 $ctrlID{$sampleid},"\t$MeanDp_hash{$ctrlID{$sampleid}}\t$tenDp_hash{$ctrlID{$sampleid}}\t$twentyCvg_hash{$ctrlID{$sampleid}}\t$fiftyDp_hash{$ctrlID{$sampleid}}\t$hundredDp_hash{$ctrlID{$sampleid}}\n";
		close INFO;
		close Q1;
		close T1;
		close T2;
		system("copy $vartypefile_hash{$sampleid} $destindir\\$sampleid\_table3.csv");
		#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
		my $t3file = "$destindir\\$sampleid\_table3.csv";
		my %num_hash;
		open my $fh,$t3file;
		open OUT,'>',"$destindir\\mutants_statistics.csv";
		print OUT "20$sampleid\t";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			#print OUT "$tmp[0]\t$tmp[1]\n";
			push @{$num_hash{$tmp[0]}},$tmp[1];
		}
		close $fh;
		system("del $destindir\\$sampleid\_table3.csv");
		#print STDERR $vartypefile_hash{$ctrlID{$sampleid}};	#
		system("copy $vartypefile_hash{$ctrlID{$sampleid}} $destindir\\$ctrlID{$sampleid}\_table3.csv");
		#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
		my $ctrlt3file = "$destindir\\$ctrlID{$sampleid}\_table3.csv";
		open my $fh,$ctrlt3file;
		#open OUT,'>',"$destindir\\mutants_statistics.csv";
		print OUT "20$ctrlID{$sampleid}\n";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			#print OUT "$tmp[0]\t$tmp[1]\n";
			push @{$num_hash{$tmp[0]}},$tmp[1];
			print OUT $tmp[0],"\t",$num_hash{$tmp[0]}[0],"\t",$tmp[1],"\n";
		}
		close $fh;
		close OUT;
		system("del $destindir\\$ctrlID{$sampleid}\_table3.csv");
		system("copy $T4file_hash{$sampleid} $destindir\\$sampleid\_table4.csv");
		my $t4file = "$destindir\\$sampleid\_table4.csv";
		open my $fh,$t4file;
		open OUT,'>',"$destindir\\mutants_hot_treat.csv";
		print OUT encode("GB2312","基因名\t变异描述\t突变比例\t致病性评估（仅供参考）\t其他信息\n");
		while(<$fh>){
			my @tmp = split /\t/,$_;
			$tmp[2] .= "%";
			@tmp = @tmp[0..$#tmp];
			my $new = join "\t",@tmp;
			print OUT $new;
		}
		close $fh;
		close OUT;
		system("del $destindir\\$sampleid\_table4.csv");
		system("copy $T6file_hash{$sampleid} $destindir\\$sampleid\_table6.csv");
		my $t6file = "$destindir\\$sampleid\_table6.csv";
		open my $fh,$t6file;
		open OUT,'>',"$destindir\\mutants_potential_treat.csv";
		print OUT encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		my @unsorted_lines;
		while(<$fh>){	#超出20行的输出前20行，以num of harm按从小到大排，不足的全输
			chomp;
			push @unsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@unsorted_lines);
		for(@sorted_lines){
			print OUT encode("GB2312",$_);
		}
		close OUT;
		system("del $destindir\\$sampleid\_table6.csv");
		system("copy $spec_T4file_hash{$sampleid} $destindir\\$sampleid\_specific_table4.csv");
		my $spect4file = "$destindir\\$sampleid\_specific_table4.csv";
		open my $fh,$spect4file;
		open OUT,'>',"$destindir\\mutants_hot_compare.csv";
		print OUT encode("GB2312","基因名\t变异描述\t体细胞突变比例\t致病性评估（仅供参考）\t其他信息\n");
		while(<$fh>){
			my @tmp = split /\t/,$_;
			$tmp[2] .= "%";
			@tmp = @tmp[0..$#tmp];
			my $new = join "\t",@tmp;
			print OUT $new;
		}
		close $fh;
		close OUT;
		system("del $destindir\\$sampleid\_specific_table4.csv");
		system("copy $spec_T6file_hash{$sampleid} $destindir\\$sampleid\_specific_table6.csv");
		my $spect6file = "$destindir\\$sampleid\_specific_table6.csv";
		open my $fh,$spect6file;
		open OUT,'>',"$destindir\\mutants_potential_compare.csv";
		print OUT encode("GB2312","基因名\t变异描述\t体细胞突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		my @unsorted_lines;
		while(<$fh>){	#超出20行的输出前20行，以num of harm按从小到大排，不足的全输
			chomp;
			push @unsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@unsorted_lines);
		for(@sorted_lines){
			print OUT encode("GB2312",$_);
		}
		close OUT;
		system("del $destindir\\$sampleid\_specific_table6.csv");
		system("copy $pharm_hash{$sampleid} $destindir\\$sampleid\_table5.csv");
		my $t5file = "$destindir\\$sampleid\_table5.csv";
		open my $fh,$t5file;
		open OUT,'>',"$destindir\\mutants_medicine.csv";
		#print OUT encode("GB2312","化疗药物\t检测基因\t检测区域\t检测结果\t结果解读\t等级\n");
		print OUT "检测基因\t检测区域\t检测结果\t结果解读\t化疗药物\t等级\n";
		while(<$fh>){
			print OUT $_;
		}
		close $fh;
		close OUT;	
		system("del $destindir\\$sampleid\_table5.csv")
	}
}
elsif ($typenum == 3){							#case-case
	my $parser = Spreadsheet::ParseXLSX->new();
	$sampleinfo=decode("GB2312",$sampleinfo);
	$sampleinfo=encode("GB2312",$sampleinfo);
	#print $sampleinfo;
	my $workbook = $parser ->parse("$sampleinfo");
	
	my($receive_date,$sampleID,$patientname,$donorname,$kiname,$age,$gender,$sampletype,$diagnose,$hospital,$sampleQ,$sequenceQ,$Q30);
	my(%receive_date_hash,%patientname_hash,%donorname_hash,%age_hash,%gender_hash,%sampletype_hash,%diagnose_hash,%hospital_hash,%sampleQ_hash,%sequenceQ_hash,%Q30_hash);
	my(@sampleIDs,$InterSum);	#以后想办法以2行为单位读文件
	my(%ctrlreceive_date_hash,%ctrlname_hash,%ctrldonorname_hash,%ctrlkiname_hash,%ctrl_gender_hash,%ctrl_age_hash,%ctrldiagnose_hash,%ctrlhospital_hash,%ctrlID,%ctrltype,%ctrlQ,%ctrlseqQ,%ctrlQ30);
	for my $worksheet($workbook->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		my $cell = $worksheet->get_cell($row_min+1,$col_min);
		$receive_date = $cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+1);
		$sampleID=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min);
		$ctrlreceive_date_hash{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#对照收样日期
		my $cell = $worksheet->get_cell($row_min+2,$col_min+1);
		$ctrlID{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#对照样本编号
		my $cell = $worksheet->get_cell($row_min+1,$col_min+2);
		$patientname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+2);
		$ctrlname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+3);
		$donorname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+3);
		$ctrldonorname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照供者姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+4);
		$kiname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+4);
		$ctrlkiname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照家属姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+5);
		$age=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+5);
		$ctrl_age_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照年龄
		my $cell = $worksheet->get_cell($row_min+1,$col_min+6);
		$gender=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+6);
		$ctrl_gender_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照性别
		my $cell = $worksheet->get_cell($row_min+1,$col_min+7);
		$sampletype=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+7);
		$ctrltype{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照样本类型
		my $cell = $worksheet->get_cell($row_min+1,$col_min+9);
		$diagnose=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+9);
		$ctrldiagnose_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照诊断
		my $cell = $worksheet->get_cell($row_min+1,$col_min+11);
		$hospital=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+11);
		$ctrlhospital_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照送检医院
		my $cell = $worksheet->get_cell($row_min+1,$col_min+14);
		$sampleQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+14);
		$ctrlQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照样本质控
		my $cell = $worksheet->get_cell($row_min+1,$col_min+15);
		$sequenceQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+15);
		$ctrlseqQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照测序质控
		my $cell = $worksheet->get_cell($row_min+1,$col_min+16);
		$Q30=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+16);
		$ctrlQ30{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照Q30
		if($sampleID){
			$receive_date_hash{$sampleID}= $receive_date;
			$patientname_hash{$sampleID} = $patientname;
			$donorname_hash{$sampleID}   = $donorname;
			$age_hash{$sampleID}         = $age;
			$gender_hash{$sampleID}      = $gender;
			$sampletype_hash{$sampleID}  = $sampletype;
			$diagnose_hash{$sampleID}    = $diagnose;
			$hospital_hash{$sampleID}    = $hospital;
			$sampleQ_hash{$sampleID}     = $sampleQ;
			$sequenceQ_hash{$sampleID}   = $sequenceQ;
			$Q30_hash{$sampleID}         = $Q30;
		}
		#print $receive_date,"\t",$sampleID,"\t",$patientname,"\n";
		#print "$sampleID\t$sampletype\t$ctrlID{$sampleID}\t$ctrltype{$sampleID}\n";
		push @sampleIDs,$sampleID;
	}
	
	my $parser = Spreadsheet::ParseXLSX->new();
	$datadpth=decode("GB2312",$datadpth);
	$datadpth=encode("GB2312",$datadpth);
	
	my $workbook = $parser ->parse("$datadpth");
	my (%datasize_hash,%theodepth_hash);
	
	for my $worksheet($workbook->worksheet(0)){
		my ($sample_id,$datasize,$theodepth);
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$sample_id = $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+1);
			$datasize=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$theodepth=$cell? encode ('GB2312', $cell->value()) :'';
			if($sample_id){
				#$q30_hash{$sample_id} = $q30;
				$datasize_hash{$sample_id} = $datasize;
				$theodepth_hash{$sample_id}= $theodepth;
			}
		}
	}
	#print "$sampleID\t$ctrlID{$sampleID}\t$datasize_hash{$sampleID}\t$datasize_hash{$ctrlID{$sampleID}}\n";
	if (defined ($opts{o})){
		$outdir .='\\' unless ($outdir=~/\\$/);
	}
	if(defined ($opts{I})){
		$Interdir .='\\' unless ($Interdir=~/\\$/);
	}
	if(defined ($opts{c})){
		$caseInterdir .='\\' unless ($caseInterdir=~/\\$/);
	}
	my(%t4_hash,%t6_hash,%varNum_hash,%MeanDp_hash,%twentyCvg_hash,%tenDp_hash,%fiftyDp_hash,%hundredDp_hash);
	
	opendir my $dh,$Interdir;
	my @interfiles = readdir($dh);
	my(%vartypefile_hash,%dpcvgfile_hash,%pharm_hash,%T4file_hash,%T6file_hash);
	foreach my $file(@interfiles){
		if($file=~ /(\w+)\_(\w+)\_variant_types_number/){
			my $sampleid = $1;
			$vartypefile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_MeanDepth_Coverage/){
			my $sampleid = $1;
			$dpcvgfile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /Interface_summary.txt/){
			$InterSum = $file;
		}
		else{next;}
	}
	
	open my $fh,$InterSum;
	while(<$fh>){
		next if /^#/;
		chomp;
		my @tmp = split /\t/,$_;
		my $sampleid = $tmp[0];
		$varNum_hash{$sampleid} = $tmp[1];
		$t4_hash{$sampleid} = $tmp[2];
		$t6_hash{$sampleid} = $tmp[3];
		$MeanDp_hash{$sampleid} = $tmp[4];
		$twentyCvg_hash{$sampleid} = $tmp[5];
		$tenDp_hash{$sampleid} = $tmp[6];
		$fiftyDp_hash{$sampleid} = $tmp[7];
		$hundredDp_hash{$sampleid} = $tmp[8];
	}
	close $fh;
	
	print STDERR "argument -c must be set after -T CC\n" unless (defined $caseInterdir);
	opendir my $dh,$caseInterdir || die "can't open $caseInterdir\n";
	my(%spec_varNum_hash,%spec_t4_hash,%spec_t6_hash,%spec_T4file_hash,%spec_T6file_hash,%both_T4file_hash,%both_T6file_hash);
	my @caseinterfiles = readdir($dh);
	foreach my $file(@caseinterfiles){
		if ($file =~ /Interface_summary.txt/){
			open my $fh ,$file;
			while(<$fh>){
				next if /^#/;
				chomp;
				my @tmp = split /\t/,$_;
				my $sampleid = $tmp[0];
				$spec_varNum_hash{$sampleid} = $tmp[1];
				$spec_t4_hash{$sampleid} = $tmp[2];
				$spec_t6_hash{$sampleid} = $tmp[3];
			}
			close $fh;
		}
		if($file =~ /(\w+)\_(\w+)\_specific_T4.txt/){
			my $sampleid = $1;
			$spec_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+)\_specific_T6.txt/){
			my $sampleid = $1;
			$spec_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+\_){3}both_T4.txt/){
			my $sampleid = $1;
			$both_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+\_){3}both_T6.txt/){
			my $sampleid = $1;
			$both_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
	}
	closedir $dh;
	foreach my $sampleid(@sampleIDs){
		my $destindir = $outdir.$sampleid.'_'.$ctrlID{$sampleid}.'_case_case';
		system("mkdir $destindir") unless(-e $destindir);
		open INFO,'>',"$destindir\\"."sample_info.csv";
		open GEN,'>',"$destindir\\"."general.csv";
		open Q1,'>',"$destindir\\"."quality_1.csv";
		open T1,'>',"$destindir\\"."patients.csv";
		open T2,'>',"$destindir\\"."quality_2.csv";
		print Q1 encode("GB2312","样本编号\t样本质控\t测序质控\tQ30\n");
		print Q1 "$sampleid\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\n";
		print Q1 "$ctrlID{$sampleid}\t$ctrlQ{$sampleID}\t$ctrlseqQ{$sampleID}\t$ctrlQ30{$sampleID}\n";
		print INFO "name\n";
		$typenum = $typenum-1;
		print INFO "sample.num\t$typenum\n";
		print INFO "sample.type.1\t$sampletype_hash{$sampleid}\n";
		print INFO "sample.type.2\t$sampletype_hash{$ctrlID{$sampleid}}\n";
		print INFO "sample.code.1\t$sampleid\n";
		print INFO "sample.code.2\t$ctrlID{$sampleid}\n";
		print GEN "name\n";
		my $totaldata = $datasize_hash{$sampleid}+$datasize_hash{$ctrlID{$sampleid}};
		printf GEN "data\t%d\n",$totaldata;
		printf GEN "raw.depth.1\t%d\n",$theodepth_hash{$sampleid};
		printf GEN "raw.depth.2\t%d\n",$theodepth_hash{$ctrlID{$sampleid}};
		#print GEN "total.treat.before\t$varNum_hash{$sampleid}\n";
		#print GEN "total.control.before\t$varNum_hash{$ctrlID{$sampleid}}\n";
		#print GEN "total.after\t$spec_varNum_hash{$sampleid}\n";								#对比后热点总个数
		#print GEN "blood.before\t$t4_hash{$sampleid}\n";
		#print GEN "blood.before.potential\t$t6_hash{$sampleid}\n";
		#print GEN "blood.after\t$spec_t4_hash{$sampleid}\n";								#对比后血液病热点总个数
		#print GEN "blood.after.potential\t$spec_t6_hash{$sampleid}\n";						#对比后潜在血液病热点总个数
		printf GEN "accuracy.1\t%.2f\n",$Q30_hash{$sampleid};
		printf GEN "accuracy.2\t%.2f\n",$ctrlQ30{$sampleid};
		print GEN "depth.1\t$MeanDp_hash{$sampleid}\n";						#理论深度
		print GEN "depth.2\t$MeanDp_hash{$ctrlID{$sampleid}}\n";
		#print INFO "seq.qual.3\t5%\n";										#突变比例
		print GEN "detect.1\t$twentyCvg_hash{$sampleid}\n";					#检出率
		print GEN "detect.2\t$twentyCvg_hash{$ctrlID{$sampleid}}\n";
		print T1 encode("GB2312","样本编号\t$sampleid\t$ctrlID{$sampleid}\n");
		my %altername;
		$altername{$sampleid} = $ctrlname_hash{$sampleid} ? $ctrlname_hash{$sampleid} : $ctrlkiname_hash{$sampleid};
		print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\t$altername{$sampleid}\n";
		print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\t$ctrl_age_hash{$sampleid}\n");
		print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\t$ctrl_gender_hash{$sampleid}\n";
		print T2 encode("GB2312","样本编号\t平均测序深度\t覆盖度10x\t覆盖度20x\t覆盖度50x\t覆盖度100x\n");
		print T2 $sampleid,"\t$MeanDp_hash{$sampleid}\t$tenDp_hash{$sampleid}\t$twentyCvg_hash{$sampleid}\t$fiftyDp_hash{$sampleid}\t$hundredDp_hash{$sampleid}\n";
		print T2 $ctrlID{$sampleid},"\t$MeanDp_hash{$ctrlID{$sampleid}}\t$tenDp_hash{$ctrlID{$sampleid}}\t$twentyCvg_hash{$ctrlID{$sampleid}}\t$fiftyDp_hash{$ctrlID{$sampleid}}\t$hundredDp_hash{$ctrlID{$sampleid}}\n";
		close INFO;
		close Q1;
		close T1;
		close T2;
		system("copy $vartypefile_hash{$sampleid} $destindir\\$sampleid\_table3.csv");
		#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
		my $t3file = "$destindir\\$sampleid\_table3.csv";
		my %num_hash;
		open my $fh,$t3file;
		open OUT,'>',"$destindir\\mutants_statistics.csv";
		print OUT "20$sampleid\t";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			#print OUT "$tmp[0]\t$tmp[1]\n";
			push @{$num_hash{$tmp[0]}},$tmp[1];
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$sampleid\_table3.csv");
		
		system("copy $vartypefile_hash{$ctrlID{$sampleid}} $destindir\\$ctrlID{$sampleid}\_table3.csv");
		my $ctrlt3file = "$destindir\\$ctrlID{$sampleid}\_table3.csv";
		open my $fh,$ctrlt3file;
		print OUT "20$ctrlID{$sampleid}\n";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			push @{$num_hash{$tmp[0]}},$tmp[1];
			print OUT $tmp[0],"\t",$num_hash{$tmp[0]}[0],"\t",$tmp[1],"\n";
		}
		close $fh;
		close OUT;
		system("del $destindir\\$ctrlID{$sampleid}\_table3.csv");
		#############################mutants_common.csv###################################
		system("copy $both_T4file_hash{$sampleid} $destindir\\$sampleid\_both_table4.csv");
		my $t4bothfile = "$destindir\\$sampleid\_both_table4.csv";
		open my $fh,$t4bothfile || die "can't open $t4bothfile";
		open COMM,'>',"$destindir\\mutants_common.csv";
		print COMM encode("GB2312","基因名\t变异描述\t突变比例\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		my $countline =0;	#用于表4和表6整体共输出20行
		while(<$fh>){
			chomp;
			if($_ ne ''){
				my @tmp = split /\t/,$_;
				$tmp[2] .= "%\t";
				@tmp = @tmp[0..$#tmp];
				my $new = join "\t",@tmp;
				print COMM $new,encode("GB2312","\t是\n");
				$countline ++ if $new;	#加入一行值加1
			}
			else{
				next;
			}
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$sampleid\_both_table4.csv");
		system("copy $both_T6file_hash{$sampleid} $destindir\\$sampleid\_both_table6.csv");
		my $t6bothfile = "$destindir\\$sampleid\_both_table6.csv";
		open my $fh,$t6bothfile || die "can't open $t6bothfile\n";
		my @t6bothunsorted_lines;
		while(<$fh>){
			chomp;
			push @t6bothunsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@t6bothunsorted_lines);
		for(@sorted_lines){
			print COMM encode("GB2312",$_);
		}
		close COMM;
		system("del $destindir\\$sampleid\_both_table6.csv");
		##############################mutants_special_1.csv###########################################
		system("copy $spec_T4file_hash{$sampleid} $destindir\\$sampleid\_specific_table4.csv");
		my $t4specfile = "$destindir\\$sampleid\_specific_table4.csv";
		open my $fh,$t4specfile || die "can't open $t4specfile";
		open COMM,'>',"$destindir\\mutants_special_1.csv";
		print COMM encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		my $countline =0;	#用于表4和表6整体共输出20行
		while(<$fh>){
			chomp;
			if($_ ne ''){
				my @tmp = split /\t/,$_;
				$tmp[2] .= "%\t";
				@tmp = @tmp[0..$#tmp];
				my $new = join "\t",@tmp;
				print COMM $new,encode("GB2312","\t是\n");
				$countline ++ if $new;	#加入一行值加1
			}
			else{
				next;
			}
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$sampleid\_specific_table4.csv");
		system("copy $spec_T6file_hash{$sampleid} $destindir\\$sampleid\_specific_table6.csv");
		my $t6specfile = "$destindir\\$sampleid\_specific_table6.csv";
		open my $fh,$t6specfile || die "can't open $t6specfile\n";
		my @t6specunsorted_lines;
		while(<$fh>){
			chomp;
			push @t6specunsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@t6specunsorted_lines);
		for(@sorted_lines){
			print COMM encode("GB2312",$_);
		}
		close COMM;
		system("del $destindir\\$sampleid\_specific_table6.csv");
		################################mutants_special_2.csv########################################
		system("copy $spec_T4file_hash{$ctrlID{$sampleid}} $destindir\\$ctrlID{$sampleid}\_specific_table4.csv");
		my $t4specfile = "$destindir\\$ctrlID{$sampleid}\_specific_table4.csv";
		open my $fh,$t4specfile || die "can't open $t4specfile";
		open COMM,'>',"$destindir\\mutants_special_2.csv";
		print COMM encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		my $countline =0;	#用于表4和表6整体共输出20行
		while(<$fh>){
			chomp;
			if($_ ne ''){
				my @tmp = split /\t/,$_;
				$tmp[2] .= "%\t";
				@tmp = @tmp[0..$#tmp];
				my $new = join "\t",@tmp;
				print COMM $new,encode("GB2312","\t是\n");
				$countline ++ if $new;	#加入一行值加1
			}
			else{
				next;
			}
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$ctrlID{$sampleid}\_specific_table4.csv");
		system("copy $spec_T6file_hash{$ctrlID{$sampleid}} $destindir\\$ctrlID{$sampleid}\_specific_table6.csv");
		my $t6specfile = "$destindir\\$ctrlID{$sampleid}\_specific_table6.csv";
		open my $fh,$t6specfile || die "can't open $t6specfile\n";
		my @t6specunsorted_lines;
		while(<$fh>){
			chomp;
			push @t6specunsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@t6specunsorted_lines);
		for(@sorted_lines){
			print COMM encode("GB2312",$_);
		}
		close COMM;
		system("del $destindir\\$ctrlID{$sampleid}\_specific_table6.csv");
	}
}
elsif ($typenum == 4){							#pedigree
	my $parser = Spreadsheet::ParseXLSX->new();
	$sampleinfo=decode("GB2312",$sampleinfo);
	$sampleinfo=encode("GB2312",$sampleinfo);
	#print $sampleinfo;
	my $workbook = $parser ->parse("$sampleinfo");
	
	my($receive_date,$sampleID,$patientname,$donorname,$kiname,$age,$gender,$sampletype,$diagnose,$hospital,$sampleQ,$sequenceQ,$Q30);
	my(%receive_date_hash,%patientname_hash,%donorname_hash,%age_hash,%gender_hash,%sampletype_hash,%diagnose_hash,%hospital_hash,%sampleQ_hash,%sequenceQ_hash,%Q30_hash);
	my(@sampleIDs,$InterSum);	#以后想办法以3行为单位读文件
	my(%ctrl1receive_date_hash,%ctrl1name_hash,%ctrl1donorname_hash,%ctrl1kiname_hash,%ctrl1_gender_hash,%ctrl1_age_hash,%ctrl1diagnose_hash,%ctrl1hospital_hash,%ctrl1ID,%ctrl1type,%ctrl1Q,%ctrl1seqQ,%ctrl1Q30);
	my(%ctrl2receive_date_hash,%ctrl2name_hash,%ctrl2donorname_hash,%ctrl2kiname_hash,%ctrl2_gender_hash,%ctrl2_age_hash,%ctrl2diagnose_hash,%ctrl2hospital_hash,%ctrl2ID,%ctrl2type,%ctrl2Q,%ctrl2seqQ,%ctrl2Q30);
	for my $worksheet($workbook->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		my $cell = $worksheet->get_cell($row_min+1,$col_min);
		$receive_date = $cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+1);
		$sampleID=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min);
		$ctrl1receive_date_hash{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#父本收样日期
		my $cell = $worksheet->get_cell($row_min+3,$col_min);
		$ctrl2receive_date_hash{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#母本收样日期
		my $cell = $worksheet->get_cell($row_min+2,$col_min+1);
		$ctrl1ID{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#父本样本编号
		my $cell = $worksheet->get_cell($row_min+3,$col_min+1);
		$ctrl2ID{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#母本样本编号
		my $cell = $worksheet->get_cell($row_min+1,$col_min+2);
		$patientname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+2);
		$ctrl1name_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照姓名1
		my $cell = $worksheet->get_cell($row_min+3,$col_min+2);
		$ctrl2name_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照姓名2
		my $cell = $worksheet->get_cell($row_min+1,$col_min+3);
		$donorname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+3);
		$ctrl1donorname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照供者姓名1
		my $cell = $worksheet->get_cell($row_min+3,$col_min+3);
		$ctrl2donorname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照供者姓名2
		my $cell = $worksheet->get_cell($row_min+1,$col_min+4);
		$kiname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+4);
		$ctrl1kiname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本姓名
		my $cell = $worksheet->get_cell($row_min+3,$col_min+4);
		$ctrl2kiname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+5);
		$age=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+5);
		$ctrl1_age_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本年龄
		my $cell = $worksheet->get_cell($row_min+3,$col_min+5);
		$ctrl2_age_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本年龄
		my $cell = $worksheet->get_cell($row_min+1,$col_min+6);
		$gender=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+6);
		$ctrl1_gender_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本性别
		my $cell = $worksheet->get_cell($row_min+3,$col_min+6);
		$ctrl2_gender_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本性别
		my $cell = $worksheet->get_cell($row_min+1,$col_min+7);
		$sampletype=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+7);
		$ctrl1type{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本样本类型
		my $cell = $worksheet->get_cell($row_min+3,$col_min+7);
		$ctrl2type{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本样本类型
		my $cell = $worksheet->get_cell($row_min+1,$col_min+9);
		$diagnose=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+9);
		$ctrl1diagnose_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本诊断
		my $cell = $worksheet->get_cell($row_min+3,$col_min+9);
		$ctrl2diagnose_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本诊断
		my $cell = $worksheet->get_cell($row_min+1,$col_min+11);
		$hospital=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+11);
		$ctrl1hospital_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本送检医院
		my $cell = $worksheet->get_cell($row_min+3,$col_min+11);
		$ctrl2hospital_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本送检医院
		my $cell = $worksheet->get_cell($row_min+1,$col_min+14);
		$sampleQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+14);
		$ctrl1Q{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本样本质控
		my $cell = $worksheet->get_cell($row_min+3,$col_min+14);
		$ctrl2Q{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本样本质控
		my $cell = $worksheet->get_cell($row_min+1,$col_min+15);
		$sequenceQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+15);
		$ctrl1seqQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本测序质控
		my $cell = $worksheet->get_cell($row_min+3,$col_min+15);
		$ctrl2seqQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本测序质控
		my $cell = $worksheet->get_cell($row_min+1,$col_min+16);
		$Q30=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+16);
		$ctrl1Q30{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#父本Q30
		my $cell = $worksheet->get_cell($row_min+3,$col_min+16);
		$ctrl2Q30{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#母本Q30
		if($sampleID){
			$receive_date_hash{$sampleID}= $receive_date;
			$patientname_hash{$sampleID} = $patientname;
			$donorname_hash{$sampleID}   = $donorname;
			$age_hash{$sampleID}         = $age;
			$gender_hash{$sampleID}      = $gender;
			$sampletype_hash{$sampleID}  = $sampletype;
			$diagnose_hash{$sampleID}    = $diagnose;
			$hospital_hash{$sampleID}    = $hospital;
			$sampleQ_hash{$sampleID}     = $sampleQ;
			$sequenceQ_hash{$sampleID}   = $sequenceQ;
			$Q30_hash{$sampleID}         = $Q30;
		}
		#print $receive_date,"\t",$sampleID,"\t",$patientname,"\n";
		#print "$sampleID\t$sampletype\t$ctrlID{$sampleID}\t$ctrltype{$sampleID}\n";
		push @sampleIDs,$sampleID;
	}
	
	my $parser = Spreadsheet::ParseXLSX->new();
	$datadpth=decode("GB2312",$datadpth);
	$datadpth=encode("GB2312",$datadpth);
	
	my $workbook = $parser ->parse("$datadpth");
	my (%datasize_hash,%theodepth_hash);
	
	for my $worksheet($workbook->worksheet(0)){
		my ($sample_id,$datasize,$theodepth);
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$sample_id = $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+1);
			$datasize=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$theodepth=$cell? encode ('GB2312', $cell->value()) :'';
			if($sample_id){
				#$q30_hash{$sample_id} = $q30;
				$datasize_hash{$sample_id} = $datasize;
				$theodepth_hash{$sample_id}= $theodepth;
			}
		}
	}
	#print "$sampleID\t$ctrlID{$sampleID}\t$datasize_hash{$sampleID}\t$datasize_hash{$ctrlID{$sampleID}}\n";
	if (defined ($opts{o})){
		$outdir .='\\' unless ($outdir=~/\\$/);
	}
	if(defined ($opts{I})){
		$Interdir .='\\' unless ($Interdir=~/\\$/);
	}
	if(defined ($opts{c})){
		$caseInterdir .='\\' unless ($caseInterdir=~/\\$/);
	}
	my(%t4_hash,%t6_hash,%varNum_hash,%MeanDp_hash,%twentyCvg_hash,%tenDp_hash,%fiftyDp_hash,%hundredDp_hash);
	
	opendir my $dh,$Interdir;
	my @interfiles = readdir($dh);
	my(%vartypefile_hash,%dpcvgfile_hash,%pharm_hash,%T4file_hash,%T6file_hash);
	foreach my $file(@interfiles){
		if($file=~ /(\w+)\_(\w+)\_variant_types_number/){
			my $sampleid = $1;
			$vartypefile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_MeanDepth_Coverage/){
			my $sampleid = $1;
			$dpcvgfile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_Recalibrated.output.vcf/){
		#elsif($file=~ /(\w+)\_(\w+)\_DrugResponse.output.vcf/){	#新版本的文件名
			my $sampleid = $1;
			$pharm_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T4.txt/){
			my $sampleid = $1;
			$T4file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T6.txt/){
			my $sampleid = $1;
			$T6file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /Interface_summary.txt/){
			$InterSum = $file;
		}
		else{next;}
	}
	open my $fh,$InterSum;
	while(<$fh>){
		next if /^#/;
		chomp;
		my @tmp = split /\t/,$_;
		my $sampleid = $tmp[0];
		$varNum_hash{$sampleid} = $tmp[1];
		$t4_hash{$sampleid} = $tmp[2];
		$t6_hash{$sampleid} = $tmp[3];
		$MeanDp_hash{$sampleid} = $tmp[4];
		$twentyCvg_hash{$sampleid} = $tmp[5];
		$tenDp_hash{$sampleid} = $tmp[6];
		$fiftyDp_hash{$sampleid} = $tmp[7];
		$hundredDp_hash{$sampleid} = $tmp[8];
	}
	close $fh;
	
	print STDERR "argument -c must be set after -T P\n" unless (defined $caseInterdir);
	opendir my $dh,$caseInterdir || die "can't open $caseInterdir\n";
	my(%refine_hash,%comhet_hash,%denovo_hash);
	my @caseinterfiles = readdir($dh);
	foreach my $file(@caseinterfiles){
		if($file =~ /(\w+)\_(\w+)\.family.ComHet.refine.txt/){
			my $sampleid = $1;
			$refine_hash{$sampleid} = $caseInterdir.$file;
		}
	}
	closedir $dh;
	foreach my $sampleid(@sampleIDs){
		my $destindir = $outdir.$sampleid.'_'.$ctrl1ID{$sampleid}.'_'.$ctrl2ID{$sampleid}.'_pedigree';
		system("mkdir $destindir") unless(-e $destindir);
		open INFO,'>',"$destindir\\"."sample_info.csv";
		open GEN,'>',"$destindir\\"."general.csv";
		open Q1,'>',"$destindir\\"."quality_1.csv";
		open T1,'>',"$destindir\\"."patients.csv";
		open T2,'>',"$destindir\\"."quality_2.csv";
		print Q1 encode("GB2312","样本编号\t样本质控\t测序质控\tQ30\n");
		print Q1 "$sampleid\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\n";
		print Q1 "$ctrl1ID{$sampleid}\t$ctrl1Q{$sampleID}\t$ctrl1seqQ{$sampleID}\t$ctrl1Q30{$sampleID}\n";
		print Q1 "$ctrl2ID{$sampleid}\t$ctrl2Q{$sampleID}\t$ctrl2seqQ{$sampleID}\t$ctrl2Q30{$sampleID}\n";
		print INFO "name\n";
		$typenum = $typenum-1;
		print INFO "sample.num\t$typenum\n";
		print INFO "sample.type.offspring\t$sampletype_hash{$sampleid}\n";
		print INFO "sample.type.father\t$sampletype_hash{$ctrl1ID{$sampleid}}\n";
		print INFO "sample.type.mother\t$sampletype_hash{$ctrl2ID{$sampleid}}\n";
		print INFO "sample.code.offspring\t$sampleid\n";
		print INFO "sample.code.father\t$ctrl1ID{$sampleid}\n";
		print INFO "sample.code.mother\t$ctrl2ID{$sampleid}\n";
		print GEN "name\n";
		my $totaldata = $datasize_hash{$sampleid}+$datasize_hash{$ctrl1ID{$sampleid}};
		printf GEN "data\t%d\n",$totaldata;
		printf GEN "raw.depth.offspring\t%d\n",$theodepth_hash{$sampleid};
		printf GEN "raw.depth.father\t%d\n",$theodepth_hash{$ctrl1ID{$sampleid}};
		printf GEN "raw.depth.mother\t%d\n",$theodepth_hash{$ctrl2ID{$sampleid}};
		print GEN "total\t$varNum_hash{$sampleid}\n";
		#print GEN "total.control.before\t$varNum_hash{$ctrlID{$sampleid}}\n";
		#print GEN "total.after\t$spec_varNum_hash{$sampleid}\n";								#对比后热点总个数
		print GEN "blood\t$t4_hash{$sampleid}\n";
		print GEN "blood.potential\t$t6_hash{$sampleid}\n";
		#print GEN "blood.after\t$spec_t4_hash{$sampleid}\n";								#对比后血液病热点总个数
		#print GEN "blood.after.potential\t$spec_t6_hash{$sampleid}\n";						#对比后潜在血液病热点总个数
		printf GEN "accuracy.offspring\t%.2f\n",$Q30_hash{$sampleid};
		printf GEN "accuracy.father\t%.2f\n",$ctrl1Q30{$sampleid};
		printf GEN "accuracy.mother\t%.2f\n",$ctrl2Q30{$sampleid};
		print GEN "depth.offspring\t$MeanDp_hash{$sampleid}\n";						#理论深度
		print GEN "depth.father\t$MeanDp_hash{$ctrl1ID{$sampleid}}\n";
		print GEN "depth.mother\t$MeanDp_hash{$ctrl2ID{$sampleid}}\n";
		#print INFO "seq.qual.3\t5%\n";										#突变比例
		print GEN "detect.offspring\t$twentyCvg_hash{$sampleid}\n";					#检出率
		print GEN "detect.father\t$twentyCvg_hash{$ctrl1ID{$sampleid}}\n";
		print GEN "detect.mother\t$twentyCvg_hash{$ctrl2ID{$sampleid}}\n";
		print T1 encode("GB2312","样本编号\t$sampleid\t$ctrl1ID{$sampleid}\t$ctrl2ID{$sampleid}\n");
		#my %altername;
		#$altername{$sampleid} = $ctrl1name_hash{$sampleid} ? $ctrl1name_hash{$sampleid} : $ctrl1kiname_hash{$sampleid};
		print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\t$ctrl1kiname_hash{$sampleid}\t$ctrl2kiname_hash{$sampleid}\n";
		print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\t$ctrl1_age_hash{$sampleid}\t$ctrl2_age_hash{$sampleid}\n");
		print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\t$ctrl1_gender_hash{$sampleid}\t$ctrl2_gender_hash{$sampleid}\n";
		print T2 encode("GB2312","样本编号\t平均测序深度\t覆盖度10x\t覆盖度20x\t覆盖度50x\t覆盖度100x\n");
		print T2 $sampleid,"\t$MeanDp_hash{$sampleid}\t$tenDp_hash{$sampleid}\t$twentyCvg_hash{$sampleid}\t$fiftyDp_hash{$sampleid}\t$hundredDp_hash{$sampleid}\n";
		print T2 $ctrl1ID{$sampleid},"\t$MeanDp_hash{$ctrl1ID{$sampleid}}\t$tenDp_hash{$ctrl1ID{$sampleid}}\t$twentyCvg_hash{$ctrl1ID{$sampleid}}\t$fiftyDp_hash{$ctrl1ID{$sampleid}}\t$hundredDp_hash{$ctrl1ID{$sampleid}}\n";
		print T2 $ctrl2ID{$sampleid},"\t$MeanDp_hash{$ctrl2ID{$sampleid}}\t$tenDp_hash{$ctrl2ID{$sampleid}}\t$twentyCvg_hash{$ctrl2ID{$sampleid}}\t$fiftyDp_hash{$ctrl2ID{$sampleid}}\t$hundredDp_hash{$ctrl2ID{$sampleid}}\n";
		close INFO;
		close Q1;
		close T1;
		close T2;
		system("copy $vartypefile_hash{$sampleid} $destindir\\$sampleid\_table3.csv");
		#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
		my $t3file = "$destindir\\$sampleid\_table3.csv";
		my %num_hash;
		open my $fh,$t3file;
		open OUT,'>',"$destindir\\mutants_statistics.csv";
		print OUT "20$sampleid\t";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			#print OUT "$tmp[0]\t$tmp[1]\n";
			push @{$num_hash{$tmp[0]}},$tmp[1];
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$sampleid\_table3.csv");
		
		system("copy $vartypefile_hash{$ctrl1ID{$sampleid}} $destindir\\$ctrl1ID{$sampleid}\_table3.csv");
		my $ctrlt3file = "$destindir\\$ctrl1ID{$sampleid}\_table3.csv";
		open my $fh,$ctrlt3file;
		print OUT "20$ctrl1ID{$sampleid}\t";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			push @{$num_hash{$tmp[0]}},$tmp[1];
			#print OUT $tmp[0],"\t",$num_hash{$tmp[0]}[0],"\t",$tmp[1],"\n";
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$ctrl1ID{$sampleid}\_table3.csv");
		
		system("copy $vartypefile_hash{$ctrl2ID{$sampleid}} $destindir\\$ctrl2ID{$sampleid}\_table3.csv");
		my $ctrlt3file = "$destindir\\$ctrl2ID{$sampleid}\_table3.csv";
		open my $fh,$ctrlt3file;
		print OUT "20$ctrl2ID{$sampleid}\n";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			push @{$num_hash{$tmp[0]}},$tmp[1];
			print OUT $tmp[0],"\t",$num_hash{$tmp[0]}[0],"\t",$num_hash{$tmp[0]}[1],"\t",$tmp[1],"\n";
		}
		close $fh;
		close OUT;
		system("del $destindir\\$ctrl2ID{$sampleid}\_table3.csv");
		#############################mutants_hot.csv###################################
		system("copy $T4file_hash{$sampleid} $destindir\\$sampleid\_table4.csv");
		my $t4file = "$destindir\\$sampleid\_table4.csv";
		open my $fh,$t4file;
		open OUT,'>',"$destindir\\mutants_hot.csv";
		print OUT encode("GB2312","基因名\t变异描述\t突变比例\t致病性评估（仅供参考）\t其他信息\n");
		while(<$fh>){
			my @tmp = split /\t/,$_;
			$tmp[2] .= "%";
			@tmp = @tmp[0..$#tmp];
			my $new = join "\t",@tmp;
			print OUT $new;
		}
		close $fh;
		close OUT;
		system("del $destindir\\$sampleid\_table4.csv");
		#############################mutants_potential.csv###################################
		system("copy $T6file_hash{$sampleid} $destindir\\$sampleid\_table6.csv");
		my $t6file = "$destindir\\$sampleid\_table6.csv";
		open my $fh,$t6file;
		open OUT,'>',"$destindir\\mutants_potential.csv";
		print OUT encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		my @unsorted_lines;
		while(<$fh>){	#超出20行的输出前20行，以num of harm按从小到大排，不足的全输
			chomp;
			push @unsorted_lines,$_;
		}
		close $fh;
		my @sorted_lines = array_SORT(@unsorted_lines);
		for(@sorted_lines){
			print OUT encode("GB2312",$_);
		}
		close OUT;
		system("del $destindir\\$sampleid\_table6.csv");
		#############################mutants_medicine.csv###################################
		system("copy $pharm_hash{$sampleid} $destindir\\$sampleid\_table5.csv");
		my $t5file = "$destindir\\$sampleid\_table5.csv";
		open my $fh,$t5file;
		open OUT,'>',"$destindir\\mutants_medicine.csv";
		#print OUT encode("GB2312","化疗药物\t检测基因\t检测区域\t检测结果\t结果解读\t等级\n");
		print OUT "检测基因\t检测区域\t检测结果\t结果解读\t化疗药物\t等级\n";
		while(<$fh>){
			print OUT $_;
		}
		close $fh;
		close OUT;	
		system("del $destindir\\$sampleid\_table5.csv");
		##############################family.ComHet.refine.txt##############################
		system("copy $refine_hash{$sampleid} $destindir\\$sampleid\_pedigree_refine.csv");
		my $refine_file = "$destindir\\$sampleid\_pedigree_refine.csv";
		open my $comh_fh,'>',"$destindir\\mutants_pedigree_ch.csv";
		open my $homo_fh,'>',"$destindir\\mutants_pedigree_rh.csv";
		open my $denovo_fh,'>',"$destindir\\mutants_pedigree_dn.csv";
		open my $fh,$refine_file;
		
		print $comh_fh encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		print $homo_fh encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		print $denovo_fh encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
		my (@new_unsorted,@comhs,@homos,@denovos);
		my $head = <$fh>;
		chomp($head);
		my @heads = split /\t/,$head;
		my (%relate,@sorted_comhs,@sorted_homos,@sorted_denovos);
		for (my $i=0;$i<=$#heads;$i++){
			$relate{$heads[$i]} = $i;
			#print $heads[$i],"\t","$relate{$heads[$i]}\n";
		}
		while(<$fh>){
			chomp;
			next if /^#/;
			my @tmp = split /\t/,$_;
			next if $tmp[$relate{KG}] >=0.05 ;
			if($tmp[$relate{Gene}] =~ /^\"(.*)\"$/){	#去除一些有多个基因名注释时两侧的"
				$tmp[$relate{Gene}] = $1;
				#print $tmp[6],"\n";
				#print $tmp[$relate{Gene}],"\n";
			}
			if($tmp[$relate{AAChange}]=~ /^\"(.*)\"$/){	#去除一些有多个基因名注释时两侧的"
				$tmp[$relate{AAChange}] = $1;
				#print $tmp[8],"\n";
			}
			my @AAs = split /,/,$tmp[$relate{AAChange}];
			my %trans =(
			'.'=>'.',
			'synonymous_SNV'=>'同义突变',
			'nonsynonymous_SNV'=>'非同义突变',
			'nonframeshift_insertion'=>'非移码插入',
			'frameshift_insertion'=>'移码插入',
			'nonframeshift_deletion'=>'非移码缺失',
			'frameshift_deletion'=>'移码缺失',
			'stopgain'=>'无义突变',
			'stoploss'=>'终止密码子突变',
			'unknown'=>'未知',
			);
			$AAs[0]=~ s/(\w+)\:(.*)/$2/;	#去除AAchange中的基因名
			#print $AAs[0],"\t",$trans{$tmp[7]},"\n";
			$tmp[$relate{GT_ControlF}] =~ s/\'(\d)\/(\d)/$1\/$2/;
			$tmp[$relate{GT_ControlM}] =~ s/\'(\d)\/(\d)/$1\/$2/;
			$tmp[$relate{GT_Case}] =~ s/\'(\d)\/(\d)/$1\/$2/;
			my $genotype = $tmp[$relate{GT_ControlF}].":".$tmp[$relate{GT_ControlM}].":".$tmp[$relate{GT_Case}];
			my ($info,$otherinfo);
			if($tmp[$relate{KG}]<=0.01 && $tmp[$relate{KG}] ne '.'){
				$info = '人群中比例极低';
			}
			elsif($tmp[$relate{KG}]>=0.01 && $tmp[$relate{KG}]<=0.05){
				$info = '人群中比例低';
			}
			else{
				$info = '人群中比例未知';
			}
			my $numofharm = $tmp[$relate{Num_of_Harm}];
			$tmp[$relate{Num_of_Harm}] = '致病性预测数据库：'.$tmp[$relate{Num_of_Harm}]."\/9".'；'.$info;	#此处建议在1/9和人群中之间加上逗号
			$tmp[$relate{avsnp150}] = '' if $tmp[$relate{avsnp150}] eq '.';
			$tmp[$relate{COSMIC_ID}] = '' if $tmp[$relate{COSMIC_ID}] eq '.';
			#$otherinfo = $tmp[2].','.$tmp[12] if $tmp[2] && $tmp[12];
			if($tmp[$relate{avsnp150}] && $tmp[$relate{COSMIC_ID}]){	#dbSNP和COSMIC都存在，取二者，否则谁不为'.'，取谁，若都为'.'，为'.'
				$otherinfo = $tmp[$relate{avsnp150}].','.$tmp[$relate{COSMIC_ID}];
			}
			elsif($tmp[$relate{avsnp150}]){
				$otherinfo = $tmp[$relate{avsnp150}];
			}
			elsif($tmp[$relate{COSMIC_ID}]){
				$otherinfo = $tmp[$relate{COSMIC_ID}];
			}
			else{
				$otherinfo = '.';
			}
			my $AA = "$AAs[0]$trans{$tmp[$relate{ExonicFunc}]}";
			my $new = join("\t",($tmp[$relate{Gene}],$AA,$genotype,$tmp[$relate{KG}],$tmp[$relate{Num_of_Harm}],$otherinfo,$numofharm));
			push @comhs,$new if ($tmp[$relate{ComHet}] eq "Yes") && ($tmp[$relate{KG}] ne ".");	#针对输出到报告中要求KG必须有值
			push @homos,$new if ($tmp[$relate{Origin}] eq 'F&M') && ($tmp[$relate{KG}] ne ".");
			push @denovos,$new if ($tmp[$relate{Origin}] eq 'De Novo') && ($tmp[$relate{KG}] ne ".");
		}
		close $fh;
		
		@sorted_comhs = array_SORT(@comhs);
		@sorted_homos = array_SORT(@homos);
		@sorted_denovos = array_SORT(@denovos);
		
		for(@sorted_comhs){
			print $comh_fh encode("GB2312",$_);	#mutants_pedigree_ch.csv
		}
		for(@sorted_homos){
			print $homo_fh encode("GB2312",$_);	#mutants_pedigree_rh.csv
		}
		for(@sorted_denovos){
			print $denovo_fh encode("GB2312",$_);	#mutants_pedigree_dn.csv
		}
		close $comh_fh;
		close $homo_fh;
		close $denovo_fh;
		system("del $destindir\\$sampleid\_pedigree_refine.csv");
	}
}
elsif($typenum == 5){	#case-case-control
	my $parser = Spreadsheet::ParseXLSX->new();
	$sampleinfo=decode("GB2312",$sampleinfo);
	$sampleinfo=encode("GB2312",$sampleinfo);
	my $workbook = $parser ->parse("$sampleinfo");
	
	my($receive_date,$sampleID,$patientname,$donorname,$kiname,$age,$gender,$sampletype,$diagnose,$hospital,$sampleQ,$sequenceQ,$Q30);
	my(%receive_date_hash,%patientname_hash,%donorname_hash,%age_hash,%gender_hash,%sampletype_hash,%diagnose_hash,%hospital_hash,%sampleQ_hash,%sequenceQ_hash,%Q30_hash);
	my(@sampleIDs,$InterSum);
	my(%case2receive_date_hash,%case2name_hash,%case2donorname_hash,%case2kiname_hash,%case2_gender_hash,%case2_age_hash,%case2diagnose_hash,%case2hospital_hash,%case2ID,%case2type,%case2Q,%case2seqQ,%case2Q30);
	my(%ctrlreceive_date_hash,%ctrlname_hash,%ctrldonorname_hash,%ctrlkiname_hash,%ctrl_gender_hash,%ctrl_age_hash,%ctrldiagnose_hash,%ctrlhospital_hash,%ctrlID,%ctrltype,%ctrlQ,%ctrlseqQ,%ctrlQ30);
	for my $worksheet($workbook->worksheet(0)){
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		my $cell = $worksheet->get_cell($row_min+1,$col_min);
		$receive_date = $cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+1,$col_min+1);
		$sampleID=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min);
		$case2receive_date_hash{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#case2收样日期
		my $cell = $worksheet->get_cell($row_min+3,$col_min);
		$ctrlreceive_date_hash{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#ctrl收样日期
		my $cell = $worksheet->get_cell($row_min+2,$col_min+1);
		$case2ID{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#case2样本编号
		my $cell = $worksheet->get_cell($row_min+3,$col_min+1);
		$ctrlID{$sampleID} = $cell? encode ('GB2312', $cell->value()) :'';	#ctrl样本编号
		my $cell = $worksheet->get_cell($row_min+1,$col_min+2);
		$patientname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+2);
		$case2name_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2姓名
		my $cell = $worksheet->get_cell($row_min+3,$col_min+2);
		$ctrlname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';		#对照姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+3);
		$donorname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+3);
		$case2donorname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2供者姓名
		my $cell = $worksheet->get_cell($row_min+3,$col_min+3);
		$ctrldonorname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#对照供者姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+4);
		$kiname=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+4);
		$case2kiname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2家属姓名
		my $cell = $worksheet->get_cell($row_min+3,$col_min+4);
		$ctrlkiname_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl家属姓名
		my $cell = $worksheet->get_cell($row_min+1,$col_min+5);
		$age=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+5);
		$case2_age_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2年龄
		my $cell = $worksheet->get_cell($row_min+3,$col_min+5);
		$ctrl_age_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl年龄
		my $cell = $worksheet->get_cell($row_min+1,$col_min+6);
		$gender=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+6);
		$case2_gender_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2性别
		my $cell = $worksheet->get_cell($row_min+3,$col_min+6);
		$ctrl_gender_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl性别
		my $cell = $worksheet->get_cell($row_min+1,$col_min+7);
		$sampletype=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+7);
		$case2type{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2样本类型
		my $cell = $worksheet->get_cell($row_min+3,$col_min+7);
		$ctrltype{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl样本类型
		my $cell = $worksheet->get_cell($row_min+1,$col_min+9);
		$diagnose=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+9);
		$case2diagnose_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2诊断
		my $cell = $worksheet->get_cell($row_min+3,$col_min+9);
		$ctrldiagnose_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl诊断
		my $cell = $worksheet->get_cell($row_min+1,$col_min+11);
		$hospital=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+11);
		$case2hospital_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2送检医院
		my $cell = $worksheet->get_cell($row_min+3,$col_min+11);
		$ctrlhospital_hash{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl送检医院
		my $cell = $worksheet->get_cell($row_min+1,$col_min+14);
		$sampleQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+14);
		$case2Q{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2样本质控
		my $cell = $worksheet->get_cell($row_min+3,$col_min+14);
		$ctrlQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl样本质控
		my $cell = $worksheet->get_cell($row_min+1,$col_min+15);
		$sequenceQ=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+15);
		$case2seqQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2测序质控
		my $cell = $worksheet->get_cell($row_min+3,$col_min+15);
		$ctrlseqQ{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrl测序质控
		my $cell = $worksheet->get_cell($row_min+1,$col_min+16);
		$Q30=$cell? encode ('GB2312', $cell->value()) :'';
		my $cell = $worksheet->get_cell($row_min+2,$col_min+16);
		$case2Q30{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#case2Q30
		my $cell = $worksheet->get_cell($row_min+3,$col_min+16);
		$ctrlQ30{$sampleID}=$cell? encode ('GB2312', $cell->value()) :'';	#ctrlQ30
		if($sampleID){
			$receive_date_hash{$sampleID}= $receive_date;
			$patientname_hash{$sampleID} = $patientname;
			$donorname_hash{$sampleID}   = $donorname;
			$age_hash{$sampleID}         = $age;
			$gender_hash{$sampleID}      = $gender;
			$sampletype_hash{$sampleID}  = $sampletype;
			$diagnose_hash{$sampleID}    = $diagnose;
			$hospital_hash{$sampleID}    = $hospital;
			$sampleQ_hash{$sampleID}     = $sampleQ;
			$sequenceQ_hash{$sampleID}   = $sequenceQ;
			$Q30_hash{$sampleID}         = $Q30;
		}
		#print $receive_date,"\t",$sampleID,"\t",$patientname,"\n";
		#print "$sampleID\t$sampletype\t$ctrlID{$sampleID}\t$ctrltype{$sampleID}\n";
		push @sampleIDs,$sampleID;
	}
	
	my $parser = Spreadsheet::ParseXLSX->new();
	$datadpth=decode("GB2312",$datadpth);
	$datadpth=encode("GB2312",$datadpth);
	
	my $workbook = $parser ->parse("$datadpth");
	my (%datasize_hash,%theodepth_hash);
	
	for my $worksheet($workbook->worksheet(0)){
		my ($sample_id,$datasize,$theodepth);
		my ($row_min,$row_max) =$worksheet->row_range();
		my ($col_min,$col_max) =$worksheet->col_range();
		for my $row(($row_min+1) .. $row_max){
			my $cell = $worksheet->get_cell($row,$col_min);
			$sample_id = $cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+1);
			$datasize=$cell? encode ('GB2312', $cell->value()) :'';
			my $cell = $worksheet->get_cell($row,$col_min+2);
			$theodepth=$cell? encode ('GB2312', $cell->value()) :'';
			if($sample_id){
				#$q30_hash{$sample_id} = $q30;
				$datasize_hash{$sample_id} = $datasize;
				$theodepth_hash{$sample_id}= $theodepth;
			}
		}
	}
	#print "$sampleID\t$ctrlID{$sampleID}\t$datasize_hash{$sampleID}\t$datasize_hash{$ctrlID{$sampleID}}\n";
	if (defined ($opts{o})){
		$outdir .='\\' unless ($outdir=~/\\$/);
	}
	if(defined ($opts{I})){
		$Interdir .='\\' unless ($Interdir=~/\\$/);
	}
	if(defined ($opts{c})){
		$caseInterdir .='\\' unless ($caseInterdir=~/\\$/);
	}
	my(%t4_hash,%t6_hash,%varNum_hash,%MeanDp_hash,%twentyCvg_hash,%tenDp_hash,%fiftyDp_hash,%hundredDp_hash);
	
	opendir my $dh,$Interdir;
	my @interfiles = readdir($dh);
	my(%vartypefile_hash,%dpcvgfile_hash,%pharm_hash,%T4file_hash,%T6file_hash);
	foreach my $file(@interfiles){
		if($file=~ /(\w+)\_(\w+)\_variant_types_number/){
			my $sampleid = $1;
			$vartypefile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_MeanDepth_Coverage/){
			my $sampleid = $1;
			$dpcvgfile_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_Recalibrated.output.vcf/){
		#elsif($file=~ /(\w+)\_(\w+)\_DrugResponse.output.vcf/){	#新版本的文件名
			my $sampleid = $1;
			$pharm_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T4.txt/){
			my $sampleid = $1;
			$T4file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /(\w+)\_(\w+)\_T6.txt/){
			my $sampleid = $1;
			$T6file_hash{$sampleid} = $Interdir.$file;
		}
		elsif($file=~ /Interface_summary.txt/){
			$InterSum = $file;
		}
		else{next;}
	}
	open my $fh,$InterSum;
	while(<$fh>){
		next if /^#/;
		chomp;
		my @tmp = split /\t/,$_;
		my $sampleid = $tmp[0];
		$varNum_hash{$sampleid} = $tmp[1];
		$t4_hash{$sampleid} = $tmp[2];
		$t6_hash{$sampleid} = $tmp[3];
		$MeanDp_hash{$sampleid} = $tmp[4];
		$twentyCvg_hash{$sampleid} = $tmp[5];
		$tenDp_hash{$sampleid} = $tmp[6];
		$fiftyDp_hash{$sampleid} = $tmp[7];
		$hundredDp_hash{$sampleid} = $tmp[8];
	}
	close $fh;
	
	print STDERR "argument -c must be set after -T CCC\n" unless (defined $caseInterdir);
	opendir my $dh,$caseInterdir || die "can't open $caseInterdir\n";
	my(%spec_varNum_hash,%spec_t4_hash,%spec_t6_hash,%spec_T4file_hash,%spec_T6file_hash,%both_T4file_hash,%both_T6file_hash);
	my (%specspec_T4file_hash,%specspec_T6file_hash,%specspecboth_T4file_hash,%specspecboth_T6file_hash);
	my @caseinterfiles = readdir($dh);
	foreach my $file(@caseinterfiles){
		if ($file =~ /Interface_summary.txt/){
			open my $fh ,$file;
			while(<$fh>){
				next if /^#/;
				chomp;
				my @tmp = split /\t/,$_;
				my $sampleid = $tmp[0];
				$spec_varNum_hash{$sampleid} = $tmp[1];
				$spec_t4_hash{$sampleid} = $tmp[2];
				$spec_t6_hash{$sampleid} = $tmp[3];
			}
			close $fh;
		}
		###########case-case之间比较###############
		if($file =~ /(\w+)\_(\w+)\_specific_T4.txt/){
			my $sampleid = $1;
			$spec_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+)\_specific_T6.txt/){
			my $sampleid = $1;
			$spec_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+\_){3}both_T4.txt/){
			my $sampleid = $1;
			$both_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+\_){3}both_T6.txt/){
			my $sampleid = $1;
			$both_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
		###########case分别去除ctrl后比较###########
		if($file =~ /(\w+)\_(\w+)\_specific_specific_T4.txt/){
			my $sampleid = $1;
			$specspec_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+)\_specific_specific_T6.txt/){
			my $sampleid = $1;
			$specspec_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+\_){5}both_T4.txt/){
			my $sampleid = $1;
			$specspecboth_T4file_hash{$sampleid} = $caseInterdir.$file;
		}
		if($file=~ /(\w+)\_(\w+\_){5}both_T6.txt/){
			my $sampleid = $1;
			$specspecboth_T6file_hash{$sampleid} = $caseInterdir.$file;
		}
	}
	closedir $dh;
	###########################################################
	foreach my $sampleid(@sampleIDs){
		my $destindir = $outdir.$sampleid.'_'.$case2ID{$sampleid}.'_'.$ctrlID{$sampleid}.'_case_case_ctrl';
		system("mkdir $destindir") unless(-e $destindir);
		open INFO,'>',"$destindir\\"."sample_info.csv";
		open GEN,'>',"$destindir\\"."general.csv";
		open Q1,'>',"$destindir\\"."quality_1.csv";
		open T1,'>',"$destindir\\"."patinents.csv";
		open T2,'>',"$destindir\\"."quality_2.csv";
		print Q1 encode("GB2312","样本编号\t样本质控\t测序质控\tQ30\n");
		print Q1 "$sampleid\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\n";
		print Q1 "$case2ID{$sampleid}\t$case2Q{$sampleID}\t$case2seqQ{$sampleID}\t$case2Q30{$sampleID}\n";
		print Q1 "$ctrlID{$sampleid}\t$ctrlQ{$sampleID}\t$ctrlseqQ{$sampleID}\t$ctrlQ30{$sampleID}\n";
		print INFO "name\n";
		$typenum = $typenum-2;
		print INFO "sample.num\t$typenum\n";
		print INFO "sample.type.case1\t$sampletype_hash{$sampleid}\n";
		print INFO "sample.type.case2\t$sampletype_hash{$case2ID{$sampleid}}\n";
		print INFO "sample.type.ctrl\t$sampletype_hash{$ctrlID{$sampleid}}\n";
		print INFO "sample.code.case1\t$sampleid\n";
		print INFO "sample.code.case2\t$case2ID{$sampleid}\n";
		print INFO "sample.code.ctrl\t$ctrlID{$sampleid}\n";
		print GEN "name\n";
		my $totaldata = $datasize_hash{$sampleid}+$datasize_hash{$case2ID{$sampleid}}+$datasize_hash{$ctrlID{$sampleid}};
		printf GEN "data\t%d\n",$totaldata;
		printf GEN "raw.depth.case1\t%d\n",$theodepth_hash{$sampleid};
		printf GEN "raw.depth.case2\t%d\n",$theodepth_hash{$case2ID{$sampleid}};
		printf GEN "raw.depth.ctrl\t%d\n",$theodepth_hash{$ctrlID{$sampleid}};
		#print GEN "total.treat.before\t$varNum_hash{$sampleid}\n";
		#print GEN "total.control.before\t$varNum_hash{$ctrlID{$sampleid}}\n";
		#print GEN "total.after\t$spec_varNum_hash{$sampleid}\n";								#对比后热点总个数
		#print GEN "blood.before\t$t4_hash{$sampleid}\n";
		#print GEN "blood.before.potential\t$t6_hash{$sampleid}\n";
		#print GEN "blood.after\t$spec_t4_hash{$sampleid}\n";								#对比后血液病热点总个数
		#print GEN "blood.after.potential\t$spec_t6_hash{$sampleid}\n";						#对比后潜在血液病热点总个数
		printf GEN "accuracy.case1\t%.2f\n",$Q30_hash{$sampleid};
		printf GEN "accuracy.case2\t%.2f\n",$case2Q30{$sampleid};
		printf GEN "accuracy.ctrl\t%.2f\n",$ctrlQ30{$sampleid};
		print GEN "depth.case1\t$MeanDp_hash{$sampleid}\n";						#理论深度
		print GEN "depth.case2\t$MeanDp_hash{$case2ID{$sampleid}}\n";
		print GEN "depth.ctrl\t$MeanDp_hash{$ctrlID{$sampleid}}\n";
		#print INFO "seq.qual.3\t5%\n";										#突变比例
		print GEN "detect.case1\t$twentyCvg_hash{$sampleid}\n";					#检出率
		print GEN "detect.case2\t$twentyCvg_hash{$case2ID{$sampleid}}\n";
		print GEN "detect.ctrl\t$twentyCvg_hash{$ctrlID{$sampleid}}\n";
		print T1 encode("GB2312","样本编号\t$sampleid\t$case2ID{$sampleid}\t$ctrlID{$sampleid}\n");
		my (%ctrl_altername,%case2_altername);
		$ctrl_altername{$sampleid} = $ctrlname_hash{$sampleid} ? $ctrlname_hash{$sampleid} : $ctrlkiname_hash{$sampleid};
		$case2_altername{$sampleid} = $case2name_hash{$sampleid} ? $case2name_hash{$sampleid} : $case2kiname_hash{$sampleid};
		print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\t$case2_altername{$sampleid}\t$ctrl_altername{$sampleid}\n";
		print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\t$case2_age_hash{$sampleid}\t$ctrl_age_hash{$sampleid}\n");
		print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\t$case2_gender_hash{$sampleid}\t$ctrl_gender_hash{$sampleid}\n";
		print T2 encode("GB2312","样本编号\t平均测序深度\t覆盖度10x\t覆盖度20x\t覆盖度50x\t覆盖度100x\n");
		print T2 $sampleid,"\t$MeanDp_hash{$sampleid}\t$tenDp_hash{$sampleid}\t$twentyCvg_hash{$sampleid}\t$fiftyDp_hash{$sampleid}\t$hundredDp_hash{$sampleid}\n";
		print T2 $case2ID{$sampleid},"\t$MeanDp_hash{$case2ID{$sampleid}}\t$tenDp_hash{$case2ID{$sampleid}}\t$twentyCvg_hash{$case2ID{$sampleid}}\t$fiftyDp_hash{$case2ID{$sampleid}}\t$hundredDp_hash{$case2ID{$sampleid}}\n";
		print T2 $ctrlID{$sampleid},"\t$MeanDp_hash{$ctrlID{$sampleid}}\t$tenDp_hash{$ctrlID{$sampleid}}\t$twentyCvg_hash{$ctrlID{$sampleid}}\t$fiftyDp_hash{$ctrlID{$sampleid}}\t$hundredDp_hash{$ctrlID{$sampleid}}\n";
		close INFO;
		close Q1;
		close T1;
		close T2;
		system("copy $vartypefile_hash{$sampleid} $destindir\\$sampleid\_table3.csv");
		#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
		my $t3file = "$destindir\\$sampleid\_table3.csv";
		my %num_hash;
		open my $fh,$t3file;
		open OUT,'>',"$destindir\\mutants_statistics.csv";
		print OUT "20$sampleid\t";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			#print OUT "$tmp[0]\t$tmp[1]\n";
			push @{$num_hash{$tmp[0]}},$tmp[1];
		}
		close $fh;
		#close OUT;
		system("del $destindir\\$sampleid\_table3.csv");
		
		system("copy $vartypefile_hash{$case2ID{$sampleid}} $destindir\\$case2ID{$sampleid}\_table3.csv");
		my $case2t3file = "$destindir\\$case2ID{$sampleid}\_table3.csv";
		open my $fh,$case2t3file;
		print OUT "20$case2ID{$sampleid}\t";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			push @{$num_hash{$tmp[0]}},$tmp[1];
			#print OUT $tmp[0],"\t",$num_hash{$tmp[0]}[0],"\t",$tmp[1],"\n";
		}
		close $fh;
		system("del $destindir\\$case2ID{$sampleid}\_table3.csv");
		
		system("copy $vartypefile_hash{$ctrlID{$sampleid}} $destindir\\$ctrlID{$sampleid}\_table3.csv");
		my $ctrlt3file = "$destindir\\$ctrlID{$sampleid}\_table3.csv";
		open my $fh,$ctrlt3file;
		print OUT "20$ctrlID{$sampleid}\n";
		while(<$fh>){
			chomp;
			next if /^#/;
			next unless $_;
			my @tmp = split/\t/,$_;
			$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
			$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
			$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
			$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
			$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/^synonymous/;
			$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/^nonsynonymous/;
			$tmp[0] = encode("GB2312","无义突变") if $tmp[0]=~/stopgain/;
			$tmp[0] = encode("GB2312","终止密码突变") if $tmp[0]=~/stoploss/;
			$tmp[0] = encode("GB2312","非移码插入") if $tmp[0]=~/nonframeshift_insertion/;
			$tmp[0] = encode("GB2312","非移码缺失") if $tmp[0]=~/nonframeshift_deletion/;
			$tmp[0] = encode("GB2312","移码插入") if $tmp[0]=~/frameshift_insertion/;
			$tmp[0] = encode("GB2312","移码缺失") if $tmp[0]=~/frameshift_deletion/;
			$tmp[0] = encode("GB2312","潜在致病性位点") if $tmp[0]=~/^pathogenic_site/;
			$tmp[0] = encode("GB2312","COSMIC致病性位点") if $tmp[0]=~/COSMIC_pathogenic_site/;
			$tmp[0] = encode("GB2312","Clinvar致病性位点") if $tmp[0]=~/Clinvar_pathogenic_site/;
			push @{$num_hash{$tmp[0]}},$tmp[1];
			print OUT $tmp[0],"\t",$num_hash{$tmp[0]}[0],"\t",$num_hash{$tmp[0]}[1],"\t",$tmp[1],"\n";
		}
		close $fh;
		close OUT;
		system("del $destindir\\$ctrlID{$sampleid}\_table3.csv");
	#########################case-case####################################
		open SPEC,'>',"$destindir\\mutants_special_1.csv";
		system("copy $spec_T4file_hash{$sampleid} $destindir\\$sampleid\_spec_table5.csv");
		my $t4specfile = "$destindir\\$sampleid\_spec_table5.csv";
		system("copy $spec_T6file_hash{$sampleid} $destindir\\$sampleid\_specific_table5_1.csv");
		my $t6specfile = "$destindir\\$sampleid\_specific_table5_1.csv";
		my @t4t6_lines = t4_t6_merge($t4specfile,$t6specfile);
		system("del $destindir\\$sampleid\_spec_table5.csv");
		system("del $destindir\\$sampleid\_specific_table5_1.csv");
		print SPEC encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		foreach(@t4t6_lines){
			print SPEC $_;
		}
		close SPEC;
		#########mutants_special_2.csv#####
		open SPEC,'>',"$destindir\\mutants_special_2.csv";
		system("copy $spec_T4file_hash{$case2ID{$sampleid}} $destindir\\$case2ID{$sampleid}\_spec_table6.csv");
		my $t4specfile = "$destindir\\$case2ID{$sampleid}\_spec_table6.csv";
		system("copy $spec_T6file_hash{$case2ID{$sampleid}} $destindir\\$case2ID{$sampleid}\_specific_table6_1.csv");
		my $t6specfile = "$destindir\\$case2ID{$sampleid}\_specific_table6_1.csv";
		my @t4t6_lines = t4_t6_merge($t4specfile,$t6specfile);
		system("del $destindir\\$case2ID{$sampleid}\_spec_table6.csv");
		system("del $destindir\\$case2ID{$sampleid}\_specific_table6_1.csv");
		print SPEC encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		foreach(@t4t6_lines){
			print SPEC $_;
		}
		close SPEC;
		#########mutants_common_1.csv#####
		open BOTH,'>',"$destindir\\mutants_common_1.csv";
		system("copy $both_T4file_hash{$sampleid} $destindir\\$sampleid\_both_table4.csv");
		my $t4bothfile = "$destindir\\$sampleid\_both_table4.csv";
		system("copy $both_T6file_hash{$sampleid} $destindir\\$sampleid\_both_table4_1.csv");
		my $t6bothfile = "$destindir\\$sampleid\_both_table4_1.csv";
		my @t4t6_lines = t4_t6_merge($t4bothfile,$t6bothfile);
		system("del $destindir\\$sampleid\_both_table4.csv");
		system("del $destindir\\$sampleid\_both_table4_1.csv");
		print BOTH encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		foreach(@t4t6_lines){
			print BOTH $_;
		}
		close BOTH;
		#########################case-case-ctrl####################################
		#########mutants_special_3.csv#####
		open SPEC,'>',"$destindir\\mutants_special_3.csv";
		system("copy $specspec_T4file_hash{$sampleid} $destindir\\$sampleid\_spec_spec_table8.csv");
		my $t4specfile = "$destindir\\$sampleid\_spec_spec_table8.csv";
		system("copy $specspec_T6file_hash{$sampleid} $destindir\\$sampleid\_spec_spec_table8_1.csv");
		my $t6specfile = "$destindir\\$sampleid\_spec_spec_table8_1.csv";
		my @t4t6_lines = t4_t6_merge($t4specfile,$t6specfile);
		system("del $destindir\\$sampleid\_spec_spec_table8.csv");
		system("del $destindir\\$sampleid\_spec_spec_table8_1.csv");
		print SPEC encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		foreach(@t4t6_lines){
			print SPEC $_;
		}
		close SPEC;
		#########mutants_special_4.csv#####
		open SPEC,'>',"$destindir\\mutants_special_4.csv";
		system("copy $specspec_T4file_hash{$case2ID{$sampleid}} $destindir\\$case2ID{$sampleid}\_spec_spec_table9.csv");
		my $t4specfile = "$destindir\\$case2ID{$sampleid}\_spec_spec_table9.csv";
		system("copy $specspec_T6file_hash{$case2ID{$sampleid}} $destindir\\$case2ID{$sampleid}\_spec_spec_table9_1.csv");
		my $t6specfile = "$destindir\\$case2ID{$sampleid}\_spec_spec_table9_1.csv";
		my @t4t6_lines = t4_t6_merge($t4specfile,$t6specfile);
		system("del $destindir\\$case2ID{$sampleid}\_spec_spec_table9.csv");
		system("del $destindir\\$case2ID{$sampleid}\_spec_spec_table9_1.csv");
		print SPEC encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		foreach(@t4t6_lines){
			print SPEC $_;
		}
		close SPEC;
		#########mutants_common_2.csv#####
		open BOTH,'>',"$destindir\\mutants_common_2.csv";
		system("copy $specspecboth_T4file_hash{$sampleid} $destindir\\$sampleid\_spec_both_table7.csv");
		my $t4bothfile = "$destindir\\$sampleid\_spec_both_table7.csv";
		system("copy $specspecboth_T6file_hash{$sampleid} $destindir\\$sampleid\_spec_both_table7_1.csv");
		my $t6bothfile = "$destindir\\$sampleid\_spec_both_table7_1.csv";
		my @t4t6_lines = t4_t6_merge($t4bothfile,$t6bothfile);
		system("del $destindir\\$sampleid\_spec_both_table7.csv");
		system("del $destindir\\$sampleid\_spec_both_table7_1.csv");
		print BOTH encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\t是否热点突变\n");
		foreach(@t4t6_lines){
			print BOTH $_;
		}
		close BOTH;
	}
}
sub array_SORT{		#单纯用于排序
	my @temp = map {[$_, split /\t/]} @_;
		
	@temp = sort {
		my @a_field = @$a[1..$#$a];
		my @b_field = @$b[1..$#$b];
		$b_field[-1] <=> $a_field[-1];
	}@temp;
	
	my @sorted_lines = map {$_->[0]} @temp;
	my @final;
	if(@sorted_lines>=20){
		for (0..19){
			my $line = $sorted_lines[$_];
			my @tmp = split /\t/,$line;
			@tmp = @tmp[0..($#tmp-1)];
			my $new = join "\t",@tmp;
			push @final,$new."\n";
		}
		return @final;
	}
	else{
		foreach(@sorted_lines){
			my $line = $_;
			my @tmp = split /\t/,$line;
			@tmp = @tmp[0..($#tmp-1)];
			my $new = join "\t",@tmp;
			push @final,$new."\n";
		}
		return @final;
	}
}

sub array_SORT2{	#用于表6排序，保证输出的表4和表6行数加起来为20
	my $countline = shift @_;	#表示表4的行数
	my @temp = map {[$_, split /\t/]} @_;
	@temp = sort {
		my @a_field = @$a[1..$#$a];
		my @b_field = @$b[1..$#$b];
		$b_field[-1] <=> $a_field[-1];
	}@temp;
	my @t6lines = @_;
	my @sorted_lines = map {$_->[0]} @temp;
	my @final;
	my $final_countline = 19 - $countline;	#补足20行
	if(@t6lines >= $final_countline){
		for (0..$final_countline){
			my $line = $t6lines[$_];
			my @tmp = split /\t/,$line;
			$tmp[2] .= "%";
			@tmp = @tmp[0..($#tmp-1)];
			my $new = join "\t",@tmp;
			push @final,$new."\n";
		}
		return @final;
	}
	else{
		foreach(@t6lines){
			my $line = $_;
			my @tmp = split /\t/,$line;
			$tmp[2] .= "%";
			@tmp = @tmp[0..($#tmp-1)];
			my $new = join "\t",@tmp;
			push @final,$new."\n";
		}
		return @final;
	}
}

sub t4_t6_merge{	#用于表4表6合并，并排序
	my ($t4,$t6) = @_;
	open my $fh,$t4;
	my $countline =0;
	my @result;
	while(<$fh>){
		chomp;
		if($_ ne ''){
			my @tmp = split /\t/,$_;
			$tmp[2] .= "%\t";
			@tmp = @tmp[0..$#tmp];
			my $new = join "\t",@tmp;
			chomp($new);
			$new = $new.encode("GB2312","\t是\n");
			push @result,$new;
			$countline ++ if $new;	#加入一行值加1
		}
		else{
			next;
		}
	}
	close $fh;
	open my $fh2,$t6;
	my @t6unsorted_lines;
	while(<$fh2>){
		chomp;
		push @t6unsorted_lines,$_;
	}
	my @sortedt6_lines = array_SORT2($countline,@t6unsorted_lines);
	for(@sortedt6_lines){
		chomp($_);
		$_ .= encode("GB2312","\t否\n");
		push @result,$_;
	}
	return @result;
}

sub print_usage{
	die "Usage\:\nperl general_process.pl [options]
	-S sampleinfo.xlsx				Required;
	-d sampleDataDepth.xlsx				Required;
	-I Interface directory				Required;
	-c case-control||case-case Interface directory
	-t type	[S|C|CC|P|CCC]				Required;
	-o format output directory			Required;
	-h help\n";
}