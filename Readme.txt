Intersection of HUNDREDS rectangles, no API!
submitted to www.planet-source-code.com/vb the 15th August 2K1, by Pietro Cecchi
Email: pietrocecchi@inwind.it

Cathegory: Graphics
Level: Intermediate

Title:Intersection HUNDREDS rectangles, no API!

Curious of how to get the intersection of hundreds rectangles almost instantly 
and with no API? Then come and see!...
Arabian and italian algorithms in play!
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
This little demo shows live intersection of 2 up to hundreds rectangles. 
It produces: A) a boolean indicating whether all rectangles intersect 
             B) the intersection rectangle itself 

The homemade method for finding the intersection rectangle (in 'IntersectHomeMade' function) 
is mine and consists of defining all (2*number of rectangles)^2 intersection points (both real
and virtual) and then filtering them to find the only 4 points belonging to all rectangles. 
A very unusual way, isn't?
In the same routine, the method for finding whether or not the 2 rectangles intersect (boolean 
value) is not fully mine but derived by a recent post on the planet: 'Knowing whether two 
objects intersect' (by Tarek Said, level Intermediate, Coding Standards, submitted august 12 2K1). 

I hope yow like the effort, have fun :)