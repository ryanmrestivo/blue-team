#-----------------------------------------------------------
# EMDMgmt.pl
# 	Perl Script to parse the EMDMgmt Registry key from SOFTWARE HIVE  
# 	Script Identifies USB Device Volume Serial Number
#	Volume Serial Number is in Decimal Format - ex usbtest_27494241129
#	
#
# References
#
# B. Reninger 
#-----------------------------------------------------------
package EMDMgmt;
use strict;

my %config = (hive          => "SOFTWARE",
              osmask        => 22,
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              version       => 2012.01);

sub getConfig{return %config}

sub getShortDescr {
	return "Get EMDMgmt key info - Identifies USB Volume Serial Number";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching EMDMgmt v.".$VERSION);
    ::rptMsg("EMDMgmt v.".$VERSION); # 2012.01 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 2012.01 [fpi] + banner
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;

	my $key_path = "Microsoft\\Windows NT\\CurrentVersion\\EMDMgmt";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("EMDMgmt");
		::rptMsg($key_path);
		::rptMsg("");
		my %apps;
		my @subkeys = $key->get_list_of_subkeys();
		if (scalar(@subkeys) > 0) {
			foreach my $s (@subkeys) {
				::rptMsg($s->get_name()." [".gmtime($s->get_timestamp())."]");
				
				my @sk = $s->get_list_of_subkeys();
				if (scalar(@sk) > 0) {
					foreach my $k (@sk) {
						my $serial = $k->get_name();
						::rptMsg("  S/N: ".$serial." [".gmtime($k->get_timestamp())."]");
						my $friendly;
						eval {
							$friendly = $k->get_value("FriendlyName")->get_data();
						};
						::rptMsg("    FriendlyName  : ".$friendly) if ($friendly ne "");
						my $parent;
						eval {
							$parent = $k->get_value("ParentIdPrefix")->get_data();
						};
						::rptMsg("    ParentIdPrefix: ".$parent) if ($parent ne "");
					}
				}
				::rptMsg("");
			}
		}
		else {
			::rptMsg($key_path." has no subkeys.");
			::logMsg($key_path." has no subkeys.");
		}
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
}
1;