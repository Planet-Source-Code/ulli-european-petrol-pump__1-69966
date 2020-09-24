Class to create an irregular shaped form from a picture. The screen shot sample shows an MM-shaped form and the ears, the black circle and the button make up the form. The remainder inside the circle does not belong to the form. 



How to:
Paste a picture into the form.

The top left pixel color is assumed to be the transparent color and it must be different from the Form BackColor (unless you size the form exactly to the picture).

All other colors will define the shape of the form. These need not necessarily be continuous, ie there may be islands surrounded by the transparent color.

Paste any controls into the visible part of the form.

Adjust cursors as you think appropriate. 

On Form Load call TrimForm to shape the form 

On Left Mouse Down call GrabForm to move the form

(complex shapes will take a little longer to create and to move)

