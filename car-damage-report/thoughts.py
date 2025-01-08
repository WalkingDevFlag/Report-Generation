# Basically what this code does is that it generates a new document for every row, 
# what we have to do is list down every instance in a single csv and then instead of 
# creating a new csv every time, we dynamically generate that document para for the next row itself

# Assume we are able to collect the data in the following manner:
'''
Class ID,Damage Type,Area
0,Dent,1500
0,Dent,1800
0,Dent,1600
0,Dent,1400
0,Dent,1300
1,Scratch,1200
1,Scratch,900
1,Scratch,2000
1,Scratch,1100
1,Scratch,950
'''

# Based on this data we write a function where it will count the number of dents and scratches, for this instance 5 scratches and 5 dents.
# Below is the format already declared in template.docx

'''
Damage Breakdown:
• Class ID [Class ID]: [Damage Type]
• Number of Instances: [Number of Instances]
• Mask Areas:
• [Damage Type] [Instance Number]: [Area]
'''

# Below is What i want my output to look like
'''
Damage Breakdown:
• Class ID 0: [Damage Type] (here it will be Dents)
• Number of Instances: [Number of Instances calculated by the code snippet]
• Mask Areas:
• Dent [Instance Number]: [Area of first Dent]
• Dent [Instance Number]: [Area of second Dent]


Damage Breakdown:
• Class ID 1: [Damage Type] (here it will be Scratches)
• Number of Instances: [Number of Instances calculated by the code snippet]
• Mask Areas:
• Scratch [Instance Number]: [Area of first scratch]
• Scratch [Instance Number]: [Area of second scratch]
'''



# Edit: I succesfully did apply it, now the issue im facing is:
'''
Car Damage Report
Introduction:
This report provides an overview of the dents and scratches identified on the vehicle. The analysis includes the classification of damage, the area affected, and visual documentation of the damages.
Damage Breakdown:
• Class ID [Class ID]: [Damage Type]
• Number of Instances: [Number of Instances]
• Mask Areas:
• [Damage Type] [Instance Number]: [Area in sq. mm]

• Class ID 0: Dent
• Number of Instances: 5
• Mask Areas:
  • Dent 1: 1500 sq. mm
  • Dent 2: 1800 sq. mm
  • Dent 3: 1600 sq. mm
  • Dent 4: 1400 sq. mm
  • Dent 5: 1300 sq. mm

• Class ID 1: Scratch
• Number of Instances: 5
• Mask Areas:
  • Scratch 6: 1200 sq. mm
  • Scratch 7: 900 sq. mm
  • Scratch 8: 2000 sq. mm
  • Scratch 9: 1100 sq. mm
  • Scratch 10: 950 sq. mm
'''
# This is how my report is looking, which is totally raw.
# The Next step is to format it into a nice looking report, where 
# my font will be Calibri (body), 12pt
# and i want the bullet point text right below the above text, like when we [shift]+[enter] for next line
# and there should be sufficient amount of distance between each Class IDs...like atleast 2 line distance
# The Main update should revolve around:
'''
Damage Breakdown:
• Class ID [Class ID]: [Damage Type]
• Number of Instances: [Number of Instances]
• Mask Areas:
• [Damage Type] [Instance Number]: [Area in sq. mm]
'''
# replacing this instead of generating below it, like it should word as its working now, but the template should be not there in the final output


# Edit: I also Achieved that, here is the current Report:
'''
Car Damage Report
Introduction:
This report provides an overview of the dents and scratches identified on the vehicle. The analysis includes the classification of damage, the area affected, and visual documentation of the damages.
Damage Breakdown:
• Class ID 0: Dent
  Number of Instances: 5
  Mask Areas:
    Dent 1: 1500 sq. mm
    Dent 2: 1800 sq. mm
    Dent 3: 1600 sq. mm
    Dent 4: 1400 sq. mm
    Dent 5: 1300 sq. mm


• Class ID 1: Scratch
  Number of Instances: 5
  Mask Areas:
    Scratch 6: 1200 sq. mm
    Scratch 7: 900 sq. mm
    Scratch 8: 2000 sq. mm
    Scratch 9: 1100 sq. mm
    Scratch 10: 950 sq. mm
'''
# Now what i have to fix is, i want to label the scratches from 1 and not let it pick up from the previous count of the damage
# so i want to start the scratch from 1 and not from 6, and this will be for all, like suppose i have Dent, Scratch and Missing, then i'll have
# [Dent 1, Dent 2, Dent 3, Dent 4, Dent 5], [Scratch 1, Scratch 2, Scratch 3, Scratch 4, Scratch 5], [Missing 1, Missing 2, Missing 3, Missing 4, Missing 5]
# and not [Dent 1, Dent 2, Dent 3, Dent 4, Dent 5], [Scratch 6, Scratch 7, Scratch 8, Scratch 9, Scratch 10], [Missing 11, Missing 12, Missing 13, Missing 14, Missing 15]

# Edit: I have fixed it, here is the current Report:
'''
Car Damage Report
Introduction:
This report provides an overview of the dents and scratches identified on the vehicle. The analysis includes the classification of damage, the area affected, and visual documentation of the damages.
Damage Breakdown:
• Class ID 0: Dent
  Number of Instances: 5
  Mask Areas:
    Dent 1: 1500 sq. mm
    Dent 2: 1800 sq. mm
    Dent 3: 1600 sq. mm
    Dent 4: 1400 sq. mm
    Dent 5: 1300 sq. mm


• Class ID 1: Scratch
  Number of Instances: 5
  Mask Areas:
    Scratch 1: 1200 sq. mm
    Scratch 2: 900 sq. mm
    Scratch 3: 2000 sq. mm
    Scratch 4: 1100 sq. mm
    Scratch 5: 950 sq. mm
'''
# Now the next change that i want is that i want the text "Damage Breakdown:" to be of size 14 and bold as well,
# like just that specific text

# Edit: I have fixed it, here is the current Report:
'''
Car Damage Report
Introduction:
This report provides an overview of the dents and scratches identified on the vehicle. The analysis includes the classification of damage, the area affected, and visual documentation of the damages.
Damage Breakdown:
• Class ID 0: Dent
  Number of Instances: 5
  Mask Areas:
    Dent 1: 1500 sq. mm
    Dent 2: 1800 sq. mm
    Dent 3: 1600 sq. mm
    Dent 4: 1400 sq. mm
    Dent 5: 1300 sq. mm


• Class ID 1: Scratch
  Number of Instances: 5
  Mask Areas:
    Scratch 1: 1200 sq. mm
    Scratch 2: 900 sq. mm
    Scratch 3: 2000 sq. mm
    Scratch 4: 1100 sq. mm
    Scratch 5: 950 sq. mm
'''