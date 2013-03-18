#!/usr/bin/perl
# rename lsm files created with AutofocusMacro for ZEN according to cellBase scheme
# Files are in $dir
# naming scheme is $naming
# The Well number xxx of prefix_Wxxx_Pyyy_Tzzz.lsm correspond to the first number
# given in the naming file
use File::Copy;
my %hash;
my $naming = 'Renaming.txt'
my $dir = '/Volumes/ellenberg/DataExchange/Andrea_for_Antonio/Boni_NUP62-NUP205_batch1_02/';
open FILE, $naming or die $!;
my $key;
while (my $line = <FILE>) {
    chomp($line);
    my @entries =split(/\s+/,$line);
    #print "@entries[0]\n";
    $hash{@entries[0]} = @entries[1];
 }
if (exists($hash{"_L1"})) {
    #print "maybe\n";
}

close FILE;
opendir(DIR,$dir) or die "cannot open directory";
@files = readdir(DIR);
close(DIR);
chdir($dir);
open LOG, ">", "renaming.log" or die $!;
for my $file (@files) {
    #print $file."\n";
    my @entries = $file =~ /(\w+)\_R(\d+)/;
    my $time;
    if (exists($hash{@entries[0]})) {
        if (@entries[1] < 10) {
           $time = "0".@entries[1];
        } else {
           $time = @entries[1];
        }
        print LOG "File renaming ".$file." into ".$hash{@entries[0]}."--T".$time.".lsm\n";
        rename( $file, $hash{@entries[0]}."--T".$time.".lsm") or die "Copy failed: $!";
    }
}
close LOG;