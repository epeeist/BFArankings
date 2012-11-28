#! /usr/bin/perl

# Reads the ranking spreadsheet for men's and women's foil and produces
# a list which can be read into Fencing Time 

use strict;

use Spreadsheet::ParseExcel;
use Getopt::Long;
use Scalar::Util qw(looks_like_number);
use POSIX qw(floor);

my $file = "";
my $gender = "";

my $result = GetOptions("file=s" => \$file,
                        "gender=s" => \$gender);

die("File not specified\n") if (! $file);
die("Gender not specifed\n") if (! $gender);

if ($gender !~ /^[mf]$/i) {
  die("Gender given as $gender, must be m or f\n");
}

my $ss = new Spreadsheet::ParseExcel;
my $book = $ss->Parse($file) ||
  die("Could not open spreadsheet $file: $!");

my $sheet = $book->{Worksheet}[0];

# The columns are Rank, Surname, Forename, Club, Year of Birth 
# another position and the total number of points.
# We use this total since the rank is interrupted for
# international fencers.

my $row = 0;
my $initialised = 0;
my $runningRank;

print "Rank,Running Rank,Surname,Forename,Points\n";
while ($row < $sheet->{MaxRow}) {
  my $rank = $sheet->{Cells}[$row][$0]->{Val};
  
  if (! $initialised) {
    if ($rank == 1) {
      $initialised = 1;
      $runningRank = 1;
    }
  }
  
  if ($initialised) {
    my $surname = $sheet->{Cells}[$row][1]->{Val};
    my $forename = $sheet->{Cells}[$row][2]->{Val};
    my $points = $sheet->{Cells}[$row][7]->{Val};   
   
    print "$rank,$runningRank,$forename,$surname,", floor($points), "\n";
  } 

  $runningRank++;
  $row++;
}
