#-----------------------------------------------------------
# spp_clients.pl
#   Access Software hive file to get list the volumes currently monitored by that have Volume Shadow Copy Service (VSS) in the SPP\Clients\{09F7EDC5-294E-4180-AF6A-FB0E6A0E9513} key
# 
# Change history
#   
#
# References
#    QCCIS white paper Reliably recovering evidential data from Volume Shadow Copies http://www.qccis.com/docs/publications/WP-VSS.pdf
# 
# copyright 2012 Corey Harrell
#-----------------------------------------------------------
package spp_clients;
use strict;

my %config = (hive          => "SOFTWARE",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20120131);

sub getConfig{return %config}
sub getShortDescr {
	return "Get SPP\Clients key contents";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching spp_clients v.".$VERSION);
    ::rptMsg("spp_clients v.".$VERSION); 
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); 
	
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
	my $key_path = "Microsoft\\WINDOWS NT\\CURRENTVERSION\\SPP\\Clients";
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		::rptMsg("SPP_Clients");
		::rptMsg($key_path);
		::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
		::rptMsg("");
		
		eval {
			my $global = $key->get_value("{09F7EDC5-294E-4180-AF6A-FB0E6A0E9513}")->get_data();
			::rptMsg("{09F7EDC5-294E-4180-AF6A-FB0E6A0E9513} = ".$global);
			::rptMsg("");
			::rptMsg("");
			::rptMsg("The value lists the volumes currently monitored by Volume Shadow Copy Service");
		};
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
}

1;
	
	