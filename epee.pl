#! /usr/bin/perl

# Reads the ranking spreadsheet for men's and women's epee and produces
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

my $offset = 1;

# Men's and women's files are slightly different

if ($gender !~ /^[mf]$/i) {
  die("Gender given as $gender, must be m or f\n");
}
else {
  $offset = 0 if ($gender eq "f");   
}

my $ss = new Spreadsheet::ParseExcel;
my $book = $ss->Parse($file) ||
  die("Could not open spreadsheet $file: $!");

my $sheet = $book->{Worksheet}[0];

# The columns are Rank, Name, Club and Points
# The mens ranking has an initial column marking the
# international status

my $row = 0;
my $initialised = 0;
my $runningRank;

print "Rank,Running Rank,Forename,Surname,Points\n";

while ($row < $sheet->{MaxRow}) {
  my $rank = $sheet->{Cells}[$row][$offset]->{Val};
  
  if (! $initialised) {
    if (looks_like_number($rank)) {
      $initialised = 1;
      $runningRank = 1
    }
  }
  
  if ($initialised) {
    my $fullName = $sheet->{Cells}[$row][1 + $offset]->{Val};

    last if ($fullName eq "");

    my $points = $sheet->{Cells}[$row][3 + $offset]->{Val};
    
    # Split the name at spaces
    
    my @name = split / /, $fullName;
    
    my $size = scalar @name;
    
    for (my $i = 0; $i < $size; $i++) {
      my $fcname = ucfirst(lc($name[$i]));
      $name[$i] = ucfirst($fcname);
    }
    
    my $forename = $name[$size - 1];
 
    pop(@name);
    
    my $surname = join(' ', @name);
   
    print "$rank,$runningRank,$forename,$surname,", floor($points), "\n";
  } 
  
  $runningRank++;
  $row++;
}
