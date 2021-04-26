#-----------------------------------------------------------
# userinfo.pl
# Plugin for Registry Ripper 
# Parse Microsoft Office UserInfo keys
#
# Change history
#   20111012 - checks for Office 2007 and 2003 UserInfo registry keys  
#
# References
#
# Journey Into IR "Why Is It What It Is" post located at http://journeyintoir.blogspot.com/2011/06/why-is-it-what-it-is.html
# ForensicArtifacts.com "UserInfo" post located at http://forensicartifacts.com/2011/06/userinfo-windows/
#
# copyright 2011 Corey Harrell (jIIr)
#-----------------------------------------------------------
package userinfo;
use strict;

my %config = (hive          => "NTUSER\.DAT",
              hasShortDescr => 1,
              hasDescr      => 0,
              hasRefs       => 0,
              osmask        => 22,
              version       => 20111012);

sub getConfig{return %config}
sub getShortDescr {
	return "Gets contents of user's UserInfo keys";	
}
sub getDescr{}
sub getRefs {}
sub getHive {return $config{hive};}
sub getVersion {return $config{version};}

my $VERSION = getVersion();

sub pluginmain {
	my $class = shift;
	my $hive = shift;
	::logMsg("Launching userinfo v.".$VERSION);
	my $reg = Parse::Win32Registry->new($hive);
	my $root_key = $reg->get_root_key;
	# Formatting for the output report
	::rptMsg("Microsoft Office 2003 - 2007 UserInfo Key Information");
	::rptMsg("");
	
	# Sets up array for the installed Office versions supported by plugin
	my @versions = ("11\.0","12\.0"); 
	
	# The For loop will obtain the UserInfo key values for each Office version supported by plugin
	foreach my $ver (@versions) {   
		my $product_path = "Software\\Microsoft\\Office\\".$ver."";
		my $user_path;
		my $key;
		my $key_path;
		
		# Determining the product information
		my $product;
		if ($ver eq "11\.0") { $product = 2003} 
			elsif ($ver eq "12\.0") {$product = 2007};
		
		# Verifies the presence of the Office product then sets the user_path variable to that version's UserInfo path
		if ($key = $root_key->get_subkey($product_path)) {
			if ($product == 2003) {
				$user_path = "Software\\Microsoft\\Office\\".$ver."\\Common\\UserInfo";
			} elsif ($product == 2007) {
				$user_path = "Software\\Microsoft\\Office\\Common\\UserInfo";
		}
		
		# Verifies the presence of the UnserInfo key then sets the key_path variable
		if ($key = $root_key->get_subkey($user_path)) {
			$key_path = $user_path;
			
			# Sets the $key variable to the registry key that's stored in the $key_path variable. Line is needed for Office 2007
			$key = $root_key->get_subkey($key_path);
			
		# Obtaining the values located in the UserInfo registry key
			::rptMsg("Office ".$product." UserInfo Information");
			::rptMsg($key_path);
			::rptMsg("LastWrite Time ".gmtime($key->get_timestamp())." (UTC)");
			::rptMsg("");
			
		# GET UserName Value--
			my $username;
			eval {$username = $key->get_value("UserName")->get_data()};
			if ($@) {::rptMsg("UserName value not found.")} else {::rptMsg("UserName value = ".$username)};
			::rptMsg("");
			
		# GET UserInitials Value--
			my $initials;
			eval {$initials = $key->get_value("UserInitials")->get_data()};
			if ($@) {::rptMsg("UserInitials value not found.")} else {::rptMsg("UserInitials value = ".$initials)};
			::rptMsg("");
		
		# GET CompanyName Value--
			my $companyname;
			eval {$companyname = $key->get_value("CompanyName")->get_data()};
			if ($@) {::rptMsg("CompanyName value not found.")} else {::rptMsg("CompanyName value = ".$companyname);};
			::rptMsg("");

		# GET Company Value--
			my $company;
			eval {$company = $key->get_value("Company")->get_data()};
			if ($@) {::rptMsg("Company value not found.")} else {::rptMsg("Company value = ".$company)};
			::rptMsg("");	
		
		
		} else { 
			# Documents the products without UserInfo registry keys
			::rptMsg("Office ".$product." ".$user_path." not found.");
			::logMsg("Office ".$product." ".$user_path." not found.");
			::rptMsg("");
		}
	} else {
		# Documents the products not present
		::rptMsg("Office ".$product." not found including the UserInfo key");
		::logMsg("Office ".$product." not found including the UserInfo key");
		::rptMsg("");
	}
}
}
1;