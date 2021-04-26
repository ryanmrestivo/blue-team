#-----------------------------------------------------------
# producttype.pl
#   Determine Windows product information
#
# Change history
#   20100325 [qar] * renamed to producttype.pl
#   20100713 [qar] * updated reference info, formatting
#   20110830 [fpi] + banner, no change to the version number
#
# References
#   http://support.microsoft.com/kb/181412
#   http://support.microsoft.com/kb/152078
#
# copyright 2010 Quantum Analytics Research, LLC
#-----------------------------------------------------------
package producttype;
use strict;
my %config = (hive          => "System",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20100325);

sub getConfig{return %config}
sub getShortDescr {
	return "Queries System hive for Windows Product info";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching producttype v.".$VERSION);
    ::rptMsg("producttype v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner

	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
	my $current;
	my $key_path = 'Select';
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		$current = $key->get_value("Current")->get_data();
		my $ccs = "ControlSet00".$current;
		my $prod_key_path = $ccs."\\Control\\ProductOptions";
		if (my $prod_key = $root_key->get_subkey($prod_key_path)) {
			::rptMsg($prod_key_path);
			::rptMsg("LastWrite = ".gmtime($prod_key->get_timestamp()));
			::rptMsg("");
			::rptMsg("Ref: http://support.microsoft.com/kb/152078");
			::rptMsg("     http://support.microsoft.com/kb/181412"); 
			::rptMsg("");
			my $type;
			eval {
				$type = $prod_key->get_value("ProductType")->get_data();
				::rptMsg("ProductType  = ".$type);
				::rptMsg("Ref: http://technet.microsoft.com/en-us/library/cc782360%28WS.10%29.aspx");
				::rptMsg("WinNT indicates a workstation.");
				::rptMsg("ServerNT indicates a standalone server.");
				::rptMsg("LanmanNT indicates a domain controller (pri/backup).");
			};
			::rptMsg("");
#-----------------------------------------------------------			
# http://technet.microsoft.com/en-us/library/cc784364(WS.10).aspx
#
# http://www.geoffchappell.com/viewer.htm?doc=studies/windows/
#        km/ntoskrnl/api/ex/exinit/productsuite.htm
#
#-----------------------------------------------------------			
			my $suite;
			eval {
				$suite = $prod_key->get_value("ProductSuite")->get_data();
				::rptMsg("ProductSuite = ".$suite);
				::rptMsg("Ref: http://technet.microsoft.com/en-us/library/cc784364%28WS.10%29.aspx");
			};
		}
		else {
			::rptMsg($prod_key_path." not found.");
		}
	}
	else {
		::rptMsg("Select key not found.");
	}
}
1