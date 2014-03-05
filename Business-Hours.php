<?php
/**
 * Template Name: Business Hours Template, No Sidebar
 *
 * Description: Twenty Twelve loves the no-sidebar look as much as
 * you do. Use this page template to remove the sidebar from any page.
 *
 * Tip: to remove the sidebar from all posts and pages simply remove
 * any active widgets from the Main Sidebar area, and the sidebar will
 * disappear everywhere.
 *
 * @package WordPress
 * @subpackage Twenty_Twelve
 * @since Twenty Twelve 1.0
 */

get_header(); ?>

	<div id="primary" class="site-content">
		<div id="content" role="main">
			<p>Template Name: Business Hours Template, No Sidebar</p>
			<?php while ( have_posts() ) : the_post(); ?>
				<?php get_template_part( 'content', 'page' ); ?>
				<?php comments_template( '', true ); ?>
			<?php endwhile; // end of the loop. ?>

<?php
    	//how does the data field represent 1 o'clock, with one or two digits?
	//how to turn 12:11 am into 00:11
	$blnBoolean = True;
	$options = array(
    	"hoursMonday" => "12:00 pm - 12:00 am"
	);
	$current_time = date( 'H', current_time( 'timestamp', 0 ) );
    	$explode_hours = explode(" ", $options['hoursMonday']);
	
	foreach ($explode_hours as $key) {
		
		if ($blnBoolean) {
		if (is_numeric($key[0]) && is_numeric($key[1]) && $key[2] == ":") {
			$hour_open = substr($key , 0, 2);
			if (($hour_open >= 01 || $hour_open <= 12) && $explode_hours[1] === 'am') {
				$hour_open = $hour_open + 12;
			}
			$blnBoolean = False;
			
		}
		} elseif (is_numeric($key[0]) && is_numeric($key[1]) && $key[2] == ":") {
			$hour_close = substr($key , 0, 2);
			if (($hour_close > 01 || $hour_close < 12) && $explode_hours[4] === 'am') {
				$hour_close = $hour_close + 12;
			}
			
		}
	}
	
	if (($current_time > $hour_open) && ($current_time < $hour_close) ) { 
		echo "The business is currently open"."<br />";
	} else {
		echo "The business is currently closed"."<br />";
	}
	echo "<br />";
	echo "Business Opens at: " . $hour_open . $explode_hours[1] . "<br />";
	echo "<br />";
	echo "Business Opens at: " . $hour_close . $explode_hours[4] . "<br />";
	echo "<br />";
	echo "Current Hour is: " . $current_time . " out of 24 hours";

	
?>

		</div><!-- #content -->
	</div><!-- #primary -->

<?php get_footer(); ?>