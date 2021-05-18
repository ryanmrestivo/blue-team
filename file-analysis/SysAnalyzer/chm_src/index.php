<?

error_reporting(0);

$fail = 0;
$defPage = "sysanalyzer.html";
$right = $_GET["r"];

if( strpos($right, '"') !== false ) $fail = 1;
if( strpos($right, '<') !== false ) $fail = 1;
if( strpos($right, '>') !== false ) $fail = 1;
if( strpos($right, "'") !== false ) $fail = 1;
if( strpos($right, '#') !== false ) $fail = 1;
if( strpos($right, '&') !== false ) $fail = 1;
if( strpos($right, '..')!== false ) $fail = 1;
if( strpos($right, '/') !== false ) $fail = 1;
if( strpos($right, '|') !== false ) $fail = 1;
if( strpos($right, '\\')!== false ) $fail = 1;
if( strpos($right, '=')!== false ) $fail = 1;
if( strpos($right, '?')!== false ) $fail = 1;
if( strpos($right, ' ')!== false ) $fail = 1;
if( strpos($right, '$')!== false ) $fail = 1;
if( strpos($right, ';')!== false ) $fail = 1;
if( strpos($right, '`')!== false ) $fail = 1;
if( strpos($right, '!')!== false ) $fail = 1;

if(!file_exists($right)) $fail=1;

/*if($fail==1){
	header('Location: http://www.goatse.info/');
	exit();
}*/

if($fail==1 || strlen($right) == 0) $right= $defPage;


?>

<!--frameset rows=150,* border=0>
	<frame src=top.html-->
	<frameset cols="200,*" border=0>
		<frame src=left.html name=left>
		<frame src="<?=$right?>" name=right> 
	</frameset>
<!--/frameset-->

