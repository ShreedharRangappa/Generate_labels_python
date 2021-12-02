from __future__ import print_function
from mailmerge import MailMerge
import math
import argparse
import configparser
import os


class Genetate_Printing_Labels_Custom():
    def __init__(self):
        self.conf = self.conf_settings()
        self.args = self.arg_parse()
        self.gen_lables_doc()

    @staticmethod
    def conf_settings():
        conf = configparser.ConfigParser()
        conf_file = './label_settings.ini'
        conf.read(conf_file)
        return conf

    def arg_parse(self):
        """
        Parse arguements to be used in cmd prompt

        """
        parser = argparse.ArgumentParser(description='Generate Labels ')

        parser.add_argument("--path_templet", dest="path_templet",  type=str,
                            help="Path to templet (MS word document)",default=self.conf.get('Label', 'path_templet'))

        parser.add_argument("--total_labels_per_sheet", dest="total_labels_per_sheet", type=int, help="Number of labels that can fit in a sheet",
                            default=self.conf.get('Label', 'total_labels_per_sheet'))

        parser.add_argument("--total_labels_required", dest="total_labels_required", type=int, help="Number of labels to print",
                            default=self.conf.get('Label', 'total_labels_required'))

        parser.add_argument("--message", dest="message", help="What to be printed on labels", default=self.conf.get('Label', 'msg'))

        parser.add_argument("--path_save", dest="path_save", default=self.conf.get('Label', 'path_save'), type=str,
                            help="Path to save result docx (MS word document)")

        return parser.parse_args()

    def gen_lables_doc(self):

        number_of_sheets = math.ceil(self.args.total_labels_required / self.args.total_labels_per_sheet)

        for i in range(number_of_sheets):
            if self.args.total_labels_required <= self.args.total_labels_per_sheet:
                self.split_and_merge(self.args.total_labels_required, i)
            else:
                self.split_and_merge(self.args.total_labels_per_sheet, i)
                self.args.total_labels_required = self.args.total_labels_required - self.args.total_labels_per_sheet

        print('Done')

    def split_and_merge(self, label_split, count):
        assert os.path.exists(self.args.path_templet), ' Templet path error'
        document = MailMerge(self.args.path_templet)
        # print(document.get_merge_fields())

        labels = []
        for i in range(self.args.total_labels_per_sheet):
            if i <= label_split:
                # labels.append(self.args.message + '-%d' % (i + count * self.args.total_labels_per_sheet))
                labels.append("  23/02/1988")
            else:
                labels.append('')

        # print(len(labels))

        document.merge(Label_00=labels[0], Label_01=labels[1], Label_02=labels[2], Label_03=labels[3], Label_04=labels[4],
                       Label_05=labels[5], Label_10=labels[6], Label_11=labels[7], Label_12=labels[8], Label_13=labels[9],
                       Label_14=labels[10], Label_15=labels[11], Label_20=labels[12], Label_21=labels[13], Label_22=labels[14],
                       Label_23=labels[15], Label_24=labels[16], Label_25=labels[17], Label_30=labels[18], Label_31=labels[19],
                       Label_32=labels[20], Label_33=labels[21], Label_34=labels[22], Label_35=labels[23], Label_40=labels[24],
                       Label_41=labels[25], Label_42=labels[26], Label_43=labels[27], Label_44=labels[28], Label_45=labels[29],
                       Label_50=labels[30], Label_51=labels[31], Label_52=labels[32], Label_53=labels[33], Label_54=labels[34],
                       Label_55=labels[35], Label_60=labels[36], Label_61=labels[37], Label_62=labels[38], Label_63=labels[39],
                       Label_64=labels[40], Label_65=labels[41], Label_70=labels[42], Label_71=labels[43], Label_72=labels[44],
                       Label_73=labels[45], Label_74=labels[46], Label_75=labels[47], Label_80=labels[48], Label_81=labels[49],
                       Label_82=labels[50], Label_83=labels[51], Label_84=labels[52], Label_85=labels[53])
        try:
            document.write('%s/%s-%d.docx' % (self.args.path_save, self.args.message, count))
        except:
            document.write('%s/%s-%d.docx' % (os.getcwd(), self.args.message, count))
            print('File save at %s' % os.getcwd())


if __name__ == '__main__':
    g = Genetate_Printing_Labels_Custom()
