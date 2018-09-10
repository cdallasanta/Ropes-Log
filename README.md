# Ropes-Log
Excel userform for inputting rope usage from a large challenge course.

On the "key" worksheet, press the "run program" button to initialize the userform. The user can select the element from a combobox, then input the number of uses each rope or cable received in a day. The date and initials fields are required to submit. When a rope is changed in the other sheets, the key worksheet and userform update automatically.

The xlsm document is here for fewing the project as a whole, and some of the main modules are available here.

Element names often have "b1" or another number in front of it (1-4), denoting which frame, or "climbing block," it is in. For example, b1rope3 is the third rope in the first climbing block. b2c1 is the first textbox for climbs in the second block.

cmb_element_change() shows the code when selecting an element using the combobox. It populates all rope and pcord names, as well as locks and remove from the tab order the blank spaces.

RunCode_Click() is the code that runs when the data is submitted. It checks each rope, and if there were climbs, adds the date, group, climbs, and initials in the last row for each rope on that element. It then resets the box, leaving the element name and initials in place.
