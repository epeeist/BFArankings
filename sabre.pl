#!/usr/bin/perl

use strict;

use HTML::TableExtract;
use LWP::Simple;
use POSIX qw(floor);
use Getopt::Long;

my $file = "";
my $gender = "";

my $result = GetOptions("file=s" => \$file,
                        "gender=s" => \$gender);

die("File not specified\n") if (! $file);
die("Gender not specifed\n") if (! $gender);

my $offset = 0;

if ($gender !~ /^[mf]$/i) {
  die("Gender given as $gender, must be m or f\n");
}

my $initialised = 0;

my $te  = new HTML::TableExtract();

open FILE, $file or die "Couldn't open file: $!"; 
my $content = join("", <FILE>); 
close FILE;

$te->parse($content);

my $runningRank;

print "Rank,Running Rank,Surname,Forename,Points\n";

foreach my $ts ($te->table_states) {
  foreach my $row ($ts->rows) {
    my @columns = @$row;
    
    my $rank = $columns[0 + $offset];
  
    if (! $initialised) {
      if ($rank == 1) {
        $initialised = 1;
	$runningRank = 1;
      }
    }
    
    if ($initialised) {
      my $surname = ucfirst(lc($columns[1 + $offset]));
      last if ($surname eq "");
      my $forename = $columns[2 + $offset];
      my $points = $columns[6 + $offset];
      
      print "$rank,$runningRank,$forename,$surname,", floor($points), "\n";
      $runningRank++;
    }
  }     
}
