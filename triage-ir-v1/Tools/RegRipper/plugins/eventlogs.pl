#-----------------------------------------------------------
# eventlogs.pl
#   Plugin for Registry Ripper; Access System hive file to get the
#   Event Log settings from System hive
# 
# Change history
#   20081219 [dwe] % created
#   20110830 [fpi] + banner, no change to the version number
#   20110830 [fpi] * bugfix, see code comments
#
# References
#   Eventlog Key: http://msdn.microsoft.com/en-us/library/aa363648(VS.85).aspx
# 
# Author: Don C. Weber, http://www.cutawaysecurity.com/blog/cutaway-security
#-----------------------------------------------------------
package eventlogs;
use strict;

my %config = (hive          => "System",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20110830);

sub getConfig{return %config}
sub getShortDescr {
	return "Gets Event Log settings from System hive";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching eventlogs v.".$VERSION);
    ::rptMsg("eventlogs v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".getHive().") ".getShortDescr()."\n"); # 20110830 [fpi] + banner
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
# First thing to do is get the ControlSet00x marked current...this is
# going to be used over and over again in plugins that access the system
# file
	my $current;
	my $key_path = 'Select';
	my $key;
	if ($key = $root_key->get_subkey($key_path)) {
		$current = $key->get_value("Current")->get_data();
		my $ccs = "ControlSet00".$current;
		my $win_path = $ccs."\\Services\\Eventlog";
		my $win;
		if ($win = $root_key->get_subkey($win_path)) {
			::rptMsg("EventLog Configuration");
			::rptMsg($win_path);
			::rptMsg("LastWrite Time ".gmtime($win->get_timestamp())." (UTC)");
			my $cn;
# 20110830 [fpi] * bugfix check existence of value before accessing it -- BEGIN
			# WAS: if ($cn = $win->get_value("ComputerName")->get_data()) {
            #          ::rptMsg("ComputerName = ".$cn);
            $cn = $win->get_value("ComputerName");
			if ($cn) {
				::rptMsg("ComputerName = ".$cn->get_data());
# 20110830 [fpi] * bugfix check existence of value before accessing it -- END
			}
			else {
				::rptMsg("ComputerName value not found.");
			}
		}
		else {
			::rptMsg($win_path." not found.");
		}

#		Cycle through each type of log
		my $logname;
		my $evpath;
		my $evlog;
		my @list_logs = $win->get_list_of_subkeys();
		foreach $logname (@list_logs){
			::rptMsg("");
			$evpath = $win_path."\\".$logname->get_name();
			if ($evlog = $root_key->get_subkey($evpath)) {
				::rptMsg("	".$logname->get_name()." EventLog");
				::rptMsg("	".$evpath);
				::rptMsg("	LastWrite Time ".gmtime($evlog->get_timestamp())." (UTC)");
				::rptMsg("	Configuration Settings");
# 20110830 [fpi] * bugfix check existence of value before accessing it -- BEGIN
                # WAS: ::rptMsg("Log location: ".$evlog->get_value('File')->get_data());
                if ( $evlog->get_value('File') ) {
                    ::rptMsg("		Log location: ".$evlog->get_value('File')->get_data());
                }
                else {
                    ::rptMsg("		Log location: n/a (missing 'File' value)");
                }
				# WAS: ::rptMsg("		Log Size: ".$evlog->get_value('MaxSize')->get_data()." Bytes");
                if ( $evlog->get_value('MaxSize') ) {
                    ::rptMsg("		Log Size: ".$evlog->get_value('MaxSize')->get_data()." Bytes");
                }
                else {
                    ::rptMsg("		Log Size: n/a (missing 'MaxSize' value)");
                }
# 20110830 [fpi] * bugfix check existence of value before accessing it -- END
				($evlog->get_value('AutoBackupLogFiles') == 0x0) ? ::rptMsg("		AutoBackupLogFiles is Disabled") : ::rptMsg("		AutoBackupLogFiles is Enabled")
			}
			else {
				::rptMsg($logname->get_name()." Event Log not found.");
			}				
		}		
		::rptMsg("");
		::rptMsg("Analysis Tips: For Event Log settings information check: http://msdn.microsoft.com/en-us/library/aa363648(VS.85).aspx");
	}
	else {
		::rptMsg($key_path." not found.");
		::logMsg($key_path." not found.");
	}
}
1;