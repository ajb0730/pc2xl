#!/usr/bin/env perl

#########
# Libraries and Flags

use strict;
use File::Basename 'basename';
use Getopt::Long;
use Data::Dumper;

#########
# Global Variables

our $VERSION = "v1.2.1";
our $SCRIPT  = basename($0);

#########
# usage( [ $message [, $code ] ] )
#
# Displays usage message and exits with an error code.
#
sub usage
{
    my $message = shift;
    my $code    = shift;

    print STDERR "\n\n$message\n\n\n" if $message;

    print STDERR << "EOHELP";

Power Church to Excel
$SCRIPT $VERSION

USAGE: $SCRIPT --help
    Displays this help message and exits.

USAGE: $SCRIPT [OPTIONS] file [ file ... ]
    file	    Name of Power Church report file to convert to Excel.

OPTIONS

    -v
    --verbose	    Print extra information about the file. Use more than
                    once to increase verbosity.

    -p pfx
    --prefix pfx    Use 'pfx' as the output prefix. Defaults to 's###_'.

    -s sep
    --separator sep Use 'sep' as the field separator. Defaults to ';'.

EOHELP

    # Default code is '1'
    exit ($code // 1);
}

#########
# Command-Line Options

my $verbose = 0;
my $help = 0;
my $prefix = 's###_';
my $sep = ';';
my $eol = "\n";

GetOptions(
    'v|verbose+' => \$verbose,
    'help' => \$help,
    'p|prefix=s' => \$prefix,
    's|separator=s' => \$sep,
);

$help = 1 unless @ARGV;

# Run Help if requested (NOTE: usage() will exit())
usage() if $help;

# Process each file in turn
foreach my $file (@ARGV) {

    # Check to see if we're probably going to be able to read the file
    if($file and -r $file) {

	$verbose and print STDERR "Attempting to open '$file' for reading.\n";

	# Open the file (or die trying)
	open(IN,'<',$file) 
	    or usage("ERROR: Unable to open file, $file, for reading: $!");
	# Read all of the file data into memory (@data)
	my @data = <IN>;
	# Close the file
	close IN;

	$verbose and print STDERR "Found ". scalar(@data) ." lines of data in '$file'.\n";

	# Process the file
	process($file,\@data);

    } else {
	# It's not a file... what is it?
	usage("ERROR: Unrecognized file, $file.");
    }
}

# Exit with a successful code
exit 0;

sub trim_leading_blanks
{
    my $data = shift;
    my $removed = 0;

    $verbose >= 2 and print STDERR "Removing blank lines: ";

    while(@$data and $data->[0] =~ /^\s*$/) {
	shift @$data;
	$removed++;
    }

    $verbose >= 2 and print STDERR "$removed removed.\n";
}


#########
# process( $filename, $data )
#
# Convert the report data found in $data into Excel format
#
# $filename - name of file that the data came from 
#             (used to create name of file to write the results to)
# $data - ref to list of data from file
#
sub process
{
    my $filename = shift;
    my $data = shift;

    # Remove all end-of-line characters
    $verbose >= 2 and print STDERR "Stripping end-of-line";
    foreach (@$data) {
	$_ =~ s/\r?\n$//;
    }
    $verbose >= 2 and print STDERR "\n";

    my $HEADER;
    my @FUNDS;

    while(1) {
	my $header = process_header($data);
	last unless $header;

	unless($HEADER) {
	    $HEADER = $header;
	    if($verbose and $verbose < 2) {
		print STDERR "Report Main Title: $HEADER->{NAME}\n";
		print STDERR "Report Sub-Title:  $HEADER->{TITLE}\n";
		print STDERR "Report Run Date:   $HEADER->{DATE}\n";
		print STDERR "Report Start Date: $HEADER->{RANGE_START}\n";
		print STDERR "Report End Date:   $HEADER->{RANGE_END}\n";
	    }
            print STDERR "Detected PowerChurch Version ". $HEADER->{VERSION} ."\n" if $HEADER->{VERSION};
	}

	print STDERR "Processing Page ". $header->{PAGE} ."\n";

	my $funds = process_data($data);
	if($funds and @$funds) {
	    $verbose and print STDERR "Found ". scalar(@$funds) ." funds on page ". $header->{PAGE} ."\n";
	    push @FUNDS, @$funds;
	}
    }

    $verbose and print STDERR "Found ". scalar(@FUNDS) ." total lines of report data in $filename.\n";

    write_output($filename,$HEADER,\@FUNDS);
}

sub process_header
{
    my $data = shift;

    my $NAME;		# Report Main Title
    my $TITLE;		# Report Sub-Title
    my $DATE;		# Date Report was run
    my $RANGE_START;	# Start of Date Range the Report Covers
    my $RANGE_END;	# End of Date Range the Report Covers

    my $version;        # Version of PowerChurch detected
    my $page;		# Current page number

    trim_leading_blanks($data);
    return undef unless @$data;

    # The first line with any data in it is the Report Main Title
    $NAME = shift @$data;
    $NAME =~ s/^\s+|\s+$//g;

    $verbose >= 2 and print STDERR "Report Main Title: $NAME\n";

    trim_leading_blanks($data);
    return undef unless @$data;

    # (Unless the Main Title contains the word 'Report', then it's the Sub-Title
    if($NAME =~ /\bReport\b/i) {
	($NAME,$TITLE) = (undef,$NAME);
	$verbose >= 2 and print STDERR "Demoting Main Title to Sub-Title\n";
    }

    # The next line is the Sub-Title, unless we've already set the Sub-Title
    $TITLE = shift @$data unless $TITLE;
    $TITLE =~ s/^\s+|\s+$//g;

    $verbose >= 2 and print STDERR "Report Sub-Title: $TITLE\n";

    trim_leading_blanks($data);
    return undef unless @$data;

    # The next line should be: 'Date Time ... Blah blah blah: Date to Date ... Page: #'
    my $line = shift @$data;
    $verbose >= 3 and print STDERR "Parsing line: '$line'\n";
    my @parse = $line =~ /(\d{2}\/\d{2}\/\d{4}\s\d{2}:\d{2}\s+(?:[AP]M))\s+\w[\w\s]+: (\d{2}\/\d{2}\/\d{4}) to (\d{2}\/\d{2}\/\d{4})\s+Page:\s*(\d+)\s*$/i;
    if(@parse) {
	$verbose >= 2 and print STDERR "Found a Version 9 Header\n";
    	($DATE,$RANGE_START,$RANGE_END,$page) = @parse;
        $version = '9';
    } else {
	@parse = $line =~ /\s+\w[\w\s]+: (\d{2}\/\d{2}\/\d{4}) to (\d{2}\/\d{2}\/\d{4})\s*$/i;
	if(@parse) {
	    $verbose >= 2 and print STDERR "Found a Version 7 Header\n";
	    ($RANGE_START,$RANGE_END) = @parse;
	    trim_leading_blanks($data);
	    return undef unless @$data;
	    $line = shift @$data;
	    @parse = $line =~ /\s*\w[\w\s]+: (\d{2}\/\d{2}\/\d{4})\s+Page:\s*(\d+)\s*$/i;
	    ($DATE,$page) = @parse;
            $version = '7';
	}
    }
    
    $verbose >= 2 and print STDERR "Report Run Date/Time: $DATE\n";
    $verbose >= 2 and print STDERR "Report Start Date: $RANGE_START\n";
    $verbose >= 2 and print STDERR "Report End Date: $RANGE_END\n";
    $verbose >= 2 and print STDERR "Page: $page\n";

    trim_leading_blanks($data);
    return undef unless @$data;

    # The next line should be the header line
    $line = shift @$data;
    unless($line =~ /^\s*fund\s*(?:\#)?\s*description\s*amount\s*$/i) {
	die "FATAL: Parsing problem, was expecting the header line.\n";
    }

    trim_leading_blanks($data);

    return { 
	NAME => $NAME, 
	TITLE => $TITLE, 
	DATE => $DATE, 
	RANGE_START => $RANGE_START, 
	RANGE_END => $RANGE_END,
	PAGE => $page,
        VERSION => $version,
    };
}

sub process_data
{
    my $data = shift;

    my @results;

    while(@$data and $data->[0] !~ /^\f/) {
	my $line = shift @$data;
	next if $line =~ /^\s*$/;

	my @parse = $line =~ /^\s*(\d{3})?\s*(.+)\s+(-?(\d+,)?\d+\.\d\d)[\s\032]*$/;
	if(@parse) {
	    my ($fund,$description,$amount) = @parse;
	    $fund =~ s/^\s+|\s+$//g;
	    $description =~ s/^\s+|\s+$//g;
	    $amount =~ s/^\s+|\s+$//g;
	    $amount =~ s/,//g;
	    push @results, [ $fund, $description, $amount ];
	    $verbose >= 2 and print STDERR "$fund\t$description\t$amount\n";
	} else {
	    $verbose and print STDERR "Failed to parse '$line'\n";
	    last;
	}
        last if $line =~ /\f$/;
    }
    if(@$data and $data->[0] =~ /\f/) {
	$data->[0] =~ s/\f//;
    }

    return \@results;
}

sub write_output
{
    my ($orig_filename,$header,$data) = @_;

    my $suffix = undef;
    my $filename;
    do {
	$filename = "xl". ($suffix?"-$suffix-":"") . basename($orig_filename);
	if($suffix) {
	    $suffix++;
	} else {
	    $suffix = 1;
	}
    } while(-f $filename);

    open(OUT,'>',$filename)
	or die "FATAL: Unable to open $filename for writing: $!.\n";
    print OUT join($sep,($prefix.'Fund#',$prefix.'Description',$prefix.'Amount')) .$eol;
    foreach my $ref (@$data) {
	my ($fund,$desc,$amnt) = @$ref;
	unless(length($fund)) {
	    ($fund,$desc) = ($desc,'');
	}
	print OUT join($sep,($prefix.$fund,$desc,$amnt)).$eol;
    }
    close OUT;

    $verbose and print STDERR "Converted $orig_filename to $filename.\n";
}

# vim: nowrap:ts=8:sts=4:et:nobackup
