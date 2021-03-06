#-----------------------------------------------------------
# clampitm.pl
#   Checks keys/values set by new version of Trojan.Clampi
#
# Change history
#   20100624 [qar] % created
#   20110830 [fpi] + banner, no change to the version number
#
# References
#   http://us.trendmicro.com/imperia/md/content/us/trendwatch/researchandanalysis/ilomo_external.pdf
#
# NOTE: This is purely a test plugin, and based solely on the reference.
#       It has not been tested on any systems that were known to be infected.
# 
# copyright 2010 Quantum Analytics Research, LLC
#-----------------------------------------------------------
package clampitm;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20100624);

sub getConfig{return %config}
sub getShortDescr {
	return "Checks for IOCs for Clampi (per Trend Micro)";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $ntuser = shift;
	::logMsg("Launching clampitm v.".$VERSION);
    ::rptMsg("clampitm v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".$config{hive}.") ".getShortDescr()."\n"); # 20110830 [fpi] + banner 
	my $reg = Parse::Win32Registry->new($ntuser);
	my $root_key = $reg->get_root_key;
	
	my $count = 0;
	
	my $key_path = 'Software\\Microsoft\\Internet Explorer\\Settings';
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("ClampiTM plugin");
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		::rptMsg("");
		
		my $tag = 1;
		my @list = qw/GatesList GID KeyE KeyM PID/;
		my @vals = $key->get_list_of_values();
		if (scalar (@vals) > 0) {
			foreach my $v (@vals) {
				my $name = $v->get_name();
				if (grep(/$name/,@list)) {
					::rptMsg(sprintf "%-10s %-30s",$name,$v->get_data());
					$tag = 0;
				}
			}
			if ($tag) {
				::rptMsg("No Clampi values found.");
			}
		}
		else {
			::rptMsg($key_path." has no values.");
		}
	}
	else {
		::rptMsg($key_path." not found.");
	}
}
1;