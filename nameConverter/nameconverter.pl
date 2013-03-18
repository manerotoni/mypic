#!/usr/bin/perl  
# nameconverter.pl
# renames lsm files created with AutofocusMacro for ZEN according to cellBase scheme
# The Well number xxx of prefix_Wxxx_Pyyy_Tzzz.lsm is compared to the first number of the naming scheme
# the name is then something--zzz.lsm

#This files contains the conversion scheme
my $naming = 'y:\DataExchange\Andrea_for_Antonio\Boni_Test_Batch01_01.txt';
#This directory contains the files to convert
my $dir = 'c:\Users\Antonio Politi\Desktop\Andrea_for_Antonio\Boni_Test_Batch01_01';

use File::Copy;
my %hash;

# create a hash correspoding the Well number and file name
open FILE, $naming or die "cannot open $naming: $!";
my $key;
while (my $line = <FILE>) {
    chomp($line);
    my @entries =split('--',$line);
    #print "@entries[0]\n";
    $hash{@entries[0]} = $line;
 }
close FILE;

# cycles through the directory and copy files to new name in same directory
opendir(DIR,$dir) or die "cannot open directory: $!";
@files = readdir(DIR);
close(DIR);
chdir($dir);
open LOG, ">", "renaming.log" or die $!;
my $time;
for my $file (@files) {
	#print $file."\n";
	my @entries = $file =~ /W(\d+)\_P\d+\_T(\d+)/;
	if (exists($hash{@entries[0]})) {
		$time = @entries[1];
		print LOG "File renaming ".$file." into ".$hash{@entries[0]}."--".$time.".lsm\n";
		copy( $file, $hash{@entries[0]}."--".$time.".lsm") or die "Copy failed: $!";
		#rename
	}
}
close LOG;