#-----------------------------------------------------------
# virut.pl
#   Plugin to detect artifacts of a Virut infection
# 
# Change history
#   20110830 [fpi] + banner, no change to the version number
#
# References
#   Symantec
#       http://www.symantec.com/security_response/writeup.jsp?
#              docid=2009-020411-2802-99&tabid=2
#
# copyright 2009 H. Carvey, keydet89@yahoo.com
#-----------------------------------------------------------
package virut;
use strict;

my %config = (hive          => "Software",
              osmask        => 22,
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              version       => 20090218);

sub getConfig{return %config}

sub getShortDescr {
	return "Detect Virut artifacts";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching virut v.".$VERSION);
    ::rptMsg("virut v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner

	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;

	my $key_path = "Microsoft\\Windows\\CurrentVersion\\Explorer";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		::rptMsg("");
		
		my $update;
		eval {
			$update = $key->get_value("UpdateHost")->get_data();
			::rptMsg("UpdateHost value detected!  Possible Virut infection!");
		};
		::rptMsg("UpdateHost value not found.") if ($@);
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
	::rptMsg("");
	::rptMsg("Also be sure to check the SYSTEM\\ControlSet00n\\Services\\SharedAccess\\");
	::rptMsg("Parameters\\FirewallPolicy\\DomainProfile\\AuthorizedApplications\\List key");
	::rptMsg("for exceptions added to the firewall; use the fw_config\.pl plugin.");
}
1;