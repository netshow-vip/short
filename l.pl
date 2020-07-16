use strict;
use warnings;
use GD::Simple;
 
# create a new image (width, height)
my $img = GD::Simple->new(200, 100);

# draw a red rectangle with blue borders
$img->bgcolor('red');
$img->fgcolor('blue');
$img->rectangle(10, 10, 50, 50); # (top_left_x, top_left_y, bottom_right_x, bottom_right_y)
                                 # ($x1, $y1, $x2, $y2)

# convert into png data
open my $out, '>', 'img.png' or die;
binmode $out;
print $out $img->png;
