#-----------------------------------------------------------
# brisv.pl
#   Plugin to detect the presence of Trojan.Brisv.A
# 
# Change history
#   20110830 [fpi] + banner, no change to the version number
#
# References
#   Symantec write-up
#       http://www.symantec.com/security_response/writeup.jsp?docid=2008-071823-1655-99
#   Info on URLAndExitCommandsEnabled value
#       http://support.microsoft.com/kb/828026
#
# copyright 2009 H. Carvey, keydet89@yahoo.com
#-----------------------------------------------------------
package brisv;
use strict;

my %config = (hive => "NTUSER\.DAT",
              osmask => 22,
              hasShortDescr => 1,
              hasDescr => 0,
              hasRefs => 0,
              version => 20090210);

sub getConfig{return %config}

sub getShortDescr {
    return "Detect artifacts of a Troj\.Brisv\.A infection";
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
    my $class = shift;
    my $hive = shift;
    ::logMsg("Launching brisv v.".$VERSION);
    ::rptMsg("brisv v.".$VERSION); # 20110830 [fpi] + banner
    ::rptMsg("(".$config{hive}.") ".getShortDescr()."\n"); # 20110830 [fpi] + banner
    my $reg = Parse::Win32Registry->new($hive);
    my $root_key = $reg->get_root_key;

    my $key_path = "Software\\Microsoft\\PIMSRV";
    my $key;
    if ($key = $root_key->get_subkey($key_path)) {
        ::rptMsg($key_path);
        ::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
        ::rptMsg("");

        my $mp_path = "Software\\Microsoft\\MediaPlayer\\Preferences";
        my $url;
        eval {
            $url = $key->get_subkey($mp_path)->get_value("URLAndExitCommandsEnabled")->get_data();
            ::rptMsg($mp_path."\\URLAndExitCommandsEnabled value set to ".$url);
        };
        # if an error occurs within the eval{} statement, do nothing
    }
    else {
        ::rptMsg($key_path." not found.");
    }
}
1;
