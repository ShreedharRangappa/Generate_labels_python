from __future__ import print_function
import mailmerge
from mailmerge import MailMerge
import math

import os

doc_path=r'LABEL.docx'


def gen_lables_doc(label,quantity):

    per_sheet_labels=9*6  # constant
    number_of_sheets=math.ceil(quantity/per_sheet_labels)

    count_labels=0
    for i in range(number_of_sheets):
        if quantity<=per_sheet_labels:
            split_and_merge(label,quantity,i)
        else:
            split_and_merge(label,per_sheet_labels,i)
            quantity=quantity-per_sheet_labels
        
    print('Done')
    
def split_and_merge(label,label_split,count):
    # label_split=real_label

    
    per_sheet_labels=(9*6)  # constant 54

    labels=[]
    for i in range(per_sheet_labels):
        if i<=label_split:
            labels.append(label+'-%d'%(i+count*per_sheet_labels))
        else:
            labels.append('')

    #print(len(labels))

    document = MailMerge(doc_path)
    #print(document.get_merge_fields())
    
    document.merge( Label_00=labels[0], Label_01=labels[1], Label_02=labels[2], Label_03=labels[3], Label_04=labels[4], Label_05=labels[5],
                    Label_10=labels[6], Label_11=labels[7], Label_12=labels[8], Label_13=labels[9], Label_14=labels[10],Label_15=labels[11],
                    Label_20=labels[12],Label_21=labels[13],Label_22=labels[14],Label_23=labels[15],Label_24=labels[16],Label_25=labels[17],
                    Label_30=labels[18],Label_31=labels[19],Label_32=labels[20],Label_33=labels[21],Label_34=labels[22],Label_35=labels[23],
                    Label_40=labels[24],Label_41=labels[25],Label_42=labels[26],Label_43=labels[27],Label_44=labels[28],Label_45=labels[29],
                    Label_50=labels[30],Label_51=labels[31],Label_52=labels[32],Label_53=labels[33],Label_54=labels[34],Label_55=labels[35],
                    Label_60=labels[36],Label_61=labels[37],Label_62=labels[38],Label_63=labels[39],Label_64=labels[40],Label_65=labels[41],
                    Label_70=labels[42],Label_71=labels[43],Label_72=labels[44],Label_73=labels[45],Label_74=labels[46],Label_75=labels[47],
                    Label_80=labels[48],Label_81=labels[49],Label_82=labels[50],Label_83=labels[51],Label_84=labels[52],Label_85=labels[53])

    document.write('%s-%d.docx'%(label,count))


gen_lables_doc('CSR2',54)
