#!/usr/bin/perl
# i.e perl general_process.pl -S sampleinfo.xlsx -q sampleQ30Data.xlsx -f Interface_summary.txt -t type -o format output directory
use strict;
use Spreadsheet::ParseExcel;
use Spreadsheet::ParseXLSX;
use Encode;
use utf8;
use Getopt::Std;

my $version = '0.3.0115';
my %opts;
getopts("S:d:I:f:t:o:h",\%opts);

my $sampleinfo = $opts{S};	#样本信息.xlsx
my $datadpth = $opts{d};	#各样本数据量理论深度.xlsx
my $Interdir = $opts{I};	#样本结果文件目录
my $InterSum = $opts{f};	#Interface summary
my $type = $opts{t};		#不同的分析模式，比如单样本，case-case等
my $outdir = $opts{o};		#保存根据不同的样本编号生成不同的格式化结果的目录
#my $totalfile = $opts{s};	#样本突变情况统计
#my $meandir = $opts{m};	#样本深度覆盖度情况结果目录


&print_usage unless (defined($sampleinfo));
&print_usage unless (defined($datadpth));
&print_usage unless (defined($InterSum));
&print_usage unless (defined($outdir));
&print_usage unless (defined($Interdir));
#&print_usage unless (defined($totalfile));
#&print_usage unless (defined($meandir));
&print_usage if (defined($opts{h}));

my $parser = Spreadsheet::ParseXLSX->new();
$sampleinfo=decode("GB2312",$sampleinfo);
$sampleinfo=encode("GB2312",$sampleinfo);
#print $sampleinfo;
my $workbook = $parser ->parse("$sampleinfo");

my($receive_date,$sampleID,$patientname,$donorname,$kiname,$age,$gender,$sampletype,$diagnose,$hospital,$sampleQ,$sequenceQ,$Q30);
my(%receive_date_hash,%patientname_hash,%donorname_hash,%age_hash,%gender_hash,%sampletype_hash,%diagnose_hash,%hospital_hash,%sampleQ_hash,%sequenceQ_hash,%Q30_hash);
my(@sampleIDs);

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

my(%t4_hash,%t6_hash,%varNum_hash,%MeanDp_hash,%twentyCvg_hash,%tenDp_hash,%fiftyDp_hash,%hundredDp_hash);
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

opendir my $dh,$Interdir;
my @interfiles = readdir($dh);
#my @variants = grep /variant_types_number/,@interfiles;
#my @mdepths  = grep /MeanDepth_Coverage/,@interfiles;
#my @t4s      = grep /T4.txt$/,@interfiles;
#my @t6s      = grep /T6.txt$/,@interfiles;

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
	else{next;}
}

foreach my $sampleid(@sampleIDs){
	my $destindir = $outdir.$sampleid;
	system("mkdir $destindir") unless(-e $destindir);
	open INFO,'>',"$destindir\\"."sample_info.csv";
	open Q1,'>',"$destindir\\"."quality_1.csv";
	open T1,'>',"$destindir\\"."table1.csv";
	open T2,'>',"$destindir\\"."table2.csv";
	print Q1 encode("GB2312","样本编号\t样本质控\t测序质控\tQ30\n");
	print Q1 "$sampleid\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\n";
	print INFO "name\t$patientname_hash{$sampleid}\n";
	print INFO "sample.num\t\n";	#此处后面需修改，为不同类型的样本数
	printf INFO "data\t%d\n",$datasize_hash{$sampleid};
	printf INFO "depth\t%d\n",$theodepth_hash{$sampleid};
	print INFO "code\t$sampleid\n";
	print INFO "total\t$varNum_hash{$sampleid}\n";		#此处后续要添加样本总变异数,故还需读入文件样本突变情况统计，设定参数-s 样本突变情况统计表
	print INFO "blood\t$t4_hash{$sampleid}\n";
	print INFO "blood.potential\t$t6_hash{$sampleid}\n";
	#printf INFO "seq.qual.1\t%.2f\n",$q30_hash{$sampleid};
	print INFO "seq.qual.2\t$MeanDp_hash{$sampleid}\n";						#理论深度
	print INFO "seq.qual.3\t5%\n";										#突变比例
	print INFO "seq.qual.4\t$twentyCvg_hash{$sampleid}\n";					#检出率
	print T1 encode("GB2312","样本编号\t$sampleid\n");
	print T1 encode("GB2312","姓名\t"),"$patientname_hash{$sampleid}\n";
	print T1 encode("GB2312","年龄\t$age_hash{$sampleid}\n");
	print T1 encode("GB2312","性别\t"),"$gender_hash{$sampleid}\n";
	print T2 encode("GB2312","样本编号\\[note\\]\t样本质控\\[note\\]\t测序质控\\[note\\]\tQ30\t平均测序深度\t覆盖度10x\t覆盖度20x\t覆盖度50x\t覆盖度100x\n");
	print T2 '20'.$sampleid,"\t$sampleQ_hash{$sampleid}\t$sequenceQ_hash{$sampleid}\t$Q30_hash{$sampleid}\t$MeanDp_hash{$sampleid}\t$tenDp_hash{$sampleid}\t$twentyCvg_hash{$sampleid}\t$fiftyDp_hash{$sampleid}\t$hundredDp_hash{$sampleid}\n";
	close INFO;
	close Q1;
	close T1;
	close T2;
	system("copy $vartypefile_hash{$sampleid} $destindir\\$sampleid\_table3.csv");
	#system("copy $dpcvgfile_hash{$sampleid} $destindir\\table3.csv");
	my $t3file = "$destindir\\$sampleid\_table3.csv";
	open my $fh,$t3file;
	open OUT,'>',"$destindir\\table3.csv";
	print OUT "x20$sampleid\n";
	while(<$fh>){
		chomp;
		next if /^#/;
		next unless $_;
		my @tmp = split/\t/,$_;
		$tmp[0] = encode("GB2312","总突变数") if $tmp[0]=~/total/;
		$tmp[0] = encode("GB2312","外显子区域") if $tmp[0]=~/exonic/;
		$tmp[0] = encode("GB2312","内含子区域") if $tmp[0]=~/intron/;
		$tmp[0] = encode("GB2312","基因间区域") if $tmp[0]=~/intergenic/;
		$tmp[0] = encode("GB2312","同义点突变") if $tmp[0]=~/synonymous/;
		$tmp[0] = encode("GB2312","错义点突变") if $tmp[0]=~/nonsynonymous/;
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
	open OUT,'>',"$destindir\\table4.csv";
	print OUT encode("GB2312","基因名\t变异描述\t突变比例\t致病性评估（仅供参考）\t其他信息\n");
	while(<$fh>){
		print OUT $_;
	}
	close $fh;
	close OUT;
	system("del $destindir\\$sampleid\_table4.csv");
	system("copy $T6file_hash{$sampleid} $destindir\\$sampleid\_table6.csv");
	my $t6file = "$destindir\\$sampleid\_table6.csv";
	open my $fh,$t6file;
	open OUT,'>',"$destindir\\table6.csv";
	print OUT encode("GB2312","基因名\t变异描述\t突变比例\tMAF\t致病性评估（仅供参考）\t其他信息\n");
	my @unsorted_lines;
	while(<$fh>){	#超出20行的输出前20行，以num of harm按从小到大排，不足的全输
		chomp;
		push @unsorted_lines,$_;
	}
	close $fh;
	my @temp = map {[$_, split /\t/]} @unsorted_lines;

	@temp = sort {
		my @a_field = @$a[1..$#$a];
		my @b_field = @$b[1..$#$b];
		$b_field[-1] <=> $a_field[-1];
	}@temp;
	
	my @sorted_lines = map {$_->[0]} @temp;
	if(@sorted_lines>=20){
		for (0..19){
			my $line = $sorted_lines[$_];
			my @tmp = split /\t/,$line;
			@tmp = @tmp[0..($#tmp-1)];
			my $new = join "\t",@tmp;
			print OUT $new,"\n";
		}
	}
	else{
		foreach(@sorted_lines){
			my $line = $_;
			my @tmp = split /\t/,$line;
			@tmp = @tmp[0..($#tmp-1)];
			my $new = join "\t",@tmp;
			print OUT $new,"\n";
		}
	}
	close OUT;
	system("del $destindir\\$sampleid\_table6.csv");
	system("copy $pharm_hash{$sampleid} $destindir\\$sampleid\_table5.csv");
	my $t5file = "$destindir\\$sampleid\_table5.csv";
	open my $fh,$t5file;
	open OUT,'>',"$destindir\\table5.csv";
	#print OUT encode("GB2312","化疗药物\t检测基因\t检测区域\t检测结果\t结果解读\t等级\n");
	print OUT "化疗药物\t检测基因\t检测区域\t检测结果\t结果解读\t等级\n";
	while(<$fh>){
		print OUT $_;
	}
	close $fh;
	close OUT;	
	system("del $destindir\\$sampleid\_table5.csv")
}
sub print_usage{
	die "Usage\:\nperl general_process.pl [options]
	-S sampleinfo.xlsx				Required;
	-d sampleDataDepth.xlsx				Required;
	-I Interface directory				Required;
	-f Interface_summary.txt			Required;
	-t type	[S|CS|CC|P]
	-o format output directory			Required;
	-h help\n";
}
