<?php 
   include("conex.phtml"); 
   $link=Conectarse(); 

$query = str_replace("·", "'", $_GET['query']);

$result = mysql_query($query) or die('[ Error: ' .mysql_error() .' ]');

echo "<table>\n";
while ($row = mysql_fetch_array($result, MYSQL_ASSOC)) {
    echo "\t<tr>\n";
    foreach ($row as $value) {
        echo "\t\t<td>$value</td>\n";
    }
    echo "\t</tr>\n";
}
echo "</table>\n";


mysql_free_result($result);
mysql_close($link);
?>