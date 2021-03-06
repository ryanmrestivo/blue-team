#-----------------------------------------------------------
# adoberdr.pl
# Plugin for Registry Ripper 
# Parse Adobe Reader MRU keys
#
# Change history
#   20100218 + added checks for versions 4.0, 5.0, 9.0
#   20091125 * modified output to make a bit more clear
#   20110830 [fpi] + added banner, no change to the version number
#   20110830 [fpi] * bugfix, check existence of 'sDI' value;
#                    added LastWrittenTime to every 'cN' key
#                    added version 10.0
#
# References
#
# Note: LastWrite times on c subkeys will all be the same,
#       as each subkey is modified as when a new entry is added
#
# copyright 2010 Quantum Analytics Research, LLC
#-----------------------------------------------------------
package adoberdr;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20110830);

sub getConfig{return %config}
sub getShortDescr {
	return "Gets user's Adobe Reader cRecentFiles values";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $ntuser = shift;
    ::logMsg("Launching adoberdr v.".$VERSION);
    ::rptMsg("adoberdr v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".$config{hive}.") ".getShortDescr()."\n"); # 20110830 [fpi] + banner
	my $reg = Parse::Win32Registry->new($ntuser);
	my $root_key = $reg->get_root_key;
# First, let's find out which version of Adobe Acrobat Reader is installed
	my $version;
	my $tag = 0;
# 20110830 [fpi] + added version "10.0"
	my @versions = ("4\.0","5\.0","6\.0","7\.0","8\.0","9\.0","10\.0");
	foreach my $ver (@versions) {		
		my $key_path = "Software\\Adobe\\Acrobat Reader\\".$ver."\\AVGeneral\\cRecentFiles";
		if (defined($root_key->get_subkey($key_path))) {
			$version = $ver;
			$tag = 1;
		}
	}
	
	if ($tag) {
		::rptMsg("Adobe Acrobat Reader version ".$version." located."); 
		my $key_path = "Software\\Adobe\\Acrobat Reader\\".$version."\\AVGeneral\\cRecentFiles";   
		my $key = $root_key->get_subkey($key_path);
		if ($key) {
			::rptMsg($key_path);
			::rptMsg("");
#			::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
			my %arkeys;
			my @subkeys = $key->get_list_of_subkeys();
			if (scalar @subkeys > 0) {
				foreach my $s (@subkeys) {
					my $num = $s->get_name();
# 20110830 [fpi] * bugfix check existence of value before accessing it -- BEGIN
					# WAS: my $data = $s->get_value('sDI')->get_data();
                    my $data = $s->get_value('sDI');
                    if ($data) {
                        $data = $data->get_data();
                    }
                    else {
                        $data = "empty (no 'sDI' values)";
                    }
# 20110830 [fpi] * bugfix check existence of value before accessing it -- END
					$num =~ s/^c//;
					$arkeys{$num}{lastwrite} = $s->get_timestamp();
					$arkeys{$num}{data} = $data;
				}
				::rptMsg("Most recent PDF opened: ".gmtime($arkeys{1}{lastwrite})." (UTC)");
				foreach my $k (sort keys %arkeys) {
# 20110830 [fpi] + added lastwrite output of every "cN" key
					# WAS: ::rptMsg("  c".$k."   ".$arkeys{$k}{data});
                    ::rptMsg("  ".gmtime($arkeys{$k}{lastwrite})." c".$k."  ".$arkeys{$k}{data});
				}
			}
			else {
				::rptMsg($key_path." has no subkeys.");
			}
		}
		else {
			::rptMsg("Could not access ".$key_path);
		}
	}
	else {
		::rptMsg("Adobe Acrobat Reader version not found.");
	}
}

1;