import os
import re
import datetime
import pymysql
from collections import namedtuple, defaultdict
import xml.etree.ElementTree as ET


class MachineInfo:
    def __init__(self):
        self.pattern = re.compile("\n\n(.*?)\s+?\n\n\n\n")
        self.__getmodelname__()
        self.__getmachinename__()

    def __getmodelname__(self):
        r_md = os.popen("wmic computersystem get model").read()
        t_md = self.pattern.search(r_md)
        if t_md:
            t_md = t_md.group(1)
            t_split = t_md.split(" ")
            if len(t_split) > 1:
                self.modelname = t_split[-1]
            else:
                self.modelname = t_md
        else:
            raise ValueError("Get Model-Name failed")

    def __getmachinename__(self):
        r_mn = os.popen("wmic computersystem get name").read()
        t_mn = self.pattern.search(r_mn)
        if t_mn:
            self.machinename = t_mn.group(1)
        else:
            raise ValueError("Get MachineName failed")


class SQLConnect():
    def __init__(self):
        self.host = '10.36.177.2'
        self.user = 'fen'
        self.passwd = '1qaz@WSX'
        self.db = 'gaia'
        self.tblremoteHDD()

    def __connect__(self):
        self.con = pymysql.connect(
            host=self.host,
            passwd=self.passwd,
            user=self.user,
            db=self.db
        )
        self.cur = self.con.cursor()

    def __disconnect__(self):
        self.con.close()

    def getscds_bymodelname(self, bios_modelname):
        # Filter SCD with GroupDescription like [Generic] and not like useless_groupdes
        # Filter WOS_PartNumber with W10OP(7A) & W10NOP64(7E)
        useless_groupdes = ["CMGE", "TENDER", "HOME & BUSINESS", "ALEXA"]
        self.__connect__()
        sql_exe1 = "SELECT PartNumber FROM tblmodelname WHERE Name like \'%{}%\'".format(bios_modelname)
        self.cur.execute(sql_exe1)
        result = self.cur.fetchall()
        assert len(result) > 0, "Get 0 ModelName"
        modelnames = [i[0] for i in result]
        sqlmdn = " or ".join(["Model_PartNumber=\'{}\'".format(mdn) for mdn in modelnames])
        sqlgroupdes = " AND ".join(["GroupDescription NOT LIKE \'%{}%\'".format(des) for des in useless_groupdes])
        sql_exe2 = (
            "SELECT SCD_NO, GroupDescription, Status, CreateDatetime FROM tblpatchcd "
            "WHERE ({sqlmdn}) AND ({sqlgroupdes}) AND (GroupDescription LIKE \'%GENERIC%\') "
            "AND SCD_NO!=\'\' AND (WOS_PartNumber=\'7A\' OR WOS_PartNumber=\'7E\') "
            "ORDER BY CreateDatetime DESC"
        ).format(sqlmdn=sqlmdn, sqlgroupdes=sqlgroupdes)
        self.cur.execute(sql_exe2)
        result_scdinfo = self.cur.fetchall()
        assert len(result_scdinfo) > 0, "Get 0 SCD list"
        self.__disconnect__()
        return result_scdinfo

    def getslircd_byscd(self, scdpn, lang, ostype):
        # Support WOS list =
        self.__connect__()
        sql_rcd = (
            "SELECT RCD_NO FROM tblkit WHERE SCD_NO=\'{scd}\' "
            "UNION "
            "SELECT RCD_NO FROM tblkit_new WHERE SCD_NO=\'{scd}\'".format(scd=scdpn)
        )
        self.cur.execute(sql_rcd)
        temp_rcd = self.cur.fetchall()
        sql_rcdno = " or ".join(["RCD_NO=\'{}\'".format(rcd[0]) for rcd in temp_rcd])
        sql_filterrcd = (
            "SELECT RCD_NO FROM tblpreloadpn WHERE WOS_PartNumber=\'{os}\' "
            "AND ({rcd})"
        ).format(os=ostype, rcd=sql_rcdno)
        self.cur.execute(sql_filterrcd)
        temp_frcd = self.cur.fetchall()
        temp_rcdno = " or ".join(["RCD_PN=\'{}\'".format(pn[0]) for pn in temp_frcd])
        sql_slircd = (
            "SELECT RSLKit_ID FROM tblrslkit WHERE Softload=\'{lang}\' "
            "AND ({rcdno})"
        ).format(lang=lang, rcdno=temp_rcdno)
        self.cur.execute(sql_slircd)
        slircdid = self.cur.fetchall()
        if len(slircdid) == 1:
            print(slircdid[0][0])
            return(slircdid[0][0])
        self.__disconnect__()

    def get_groupdescription(self):
        # Test Only
        self.__connect__()
        sql_exe1 = "SELECT DISTINCT GroupDescription FROM tblpatchcd WHERE GroupDescription LIKE \'%Generic%\'"
        self.cur.execute(sql_exe1)
        result = self.cur.fetchall()
        result = [i[0] for i in result]
        for i in result:
            print(i)
        self.__disconnect__()

    def tblremoteHDD(self):
        sql = "select * from tblremoteNAPP"
        self.__connect__()
        self.cur.execute(sql)
        result = self.cur.fetchall()
        self.__disconnect__()
        for image in result:
            if image[1] == "RCD":
                self.rcdaddress = image[2]
                self.rcduser = image[3]
                self.rcdpasswd = image[4]
                self.rcdmount = image[5]
            elif image[1] == "SCD":
                self.scdaddress = image[2]
                self.scduser = image[3]
                self.scdpasswd = image[4]
                self.scdmount = image[5]
            elif image[1] == "LPCD":
                self.lpcdaddress = image[2]
                self.lpcduser = image[3]
                self.lpcdpasswd = image[4]
                self.lpcdmount = image[5]
            elif image[1] == "PatchCD":
                self.pcdaddress = image[2]
                self.pcduser = image[3]
                self.pcdpasswd = image[4]
                self.pcdmount = image[5]
            elif image[1] == "Local":
                self.localaddress = image[2]
                self.localuser = image[3]
                self.localpasswd = image[4]
                self.localmount = image[5]

    def detailinfo(self):
        self.tblremoteHDD()
        print("rcdadress is {}".format(self.rcdaddress))
        print("rcduser is {}".format(self.rcduser))
        print("rcdmount is {}".format(self.rcdmount))
        print("scdadress is {}".format(self.scdaddress))
        print("scduser is {}".format(self.scduser))
        print("scdmount is {}".format(self.scdmount))
        print("lpcdadress is {}".format(self.lpcdaddress))
        print("lpcduser is {}".format(self.lpcduser))
        print("lpcdmount is {}".format(self.lpcdmount))
        print("pcdadress is {}".format(self.pcdaddress))
        print("pcduser is {}".format(self.pcduser))
        print("pcdmount is {}".format(self.pcdmount))
        print("localadress is {}".format(self.localaddress))
        print("localuser is {}".format(self.localuser))
        print("localmount is {}".format(self.localmount))


class RemoteHDD(SQLConnect):
    def __init__(self):
        SQLConnect.__init__(self)
        self.tblremoteHDD()
        RemoteFolder = namedtuple("RemoveFolder", "address, user, password, letter")
        self.rcd = RemoteFolder(self.rcdaddress, self.rcduser, self.rcdpasswd, self.rcdmount)
        self.scd = RemoteFolder(self.scdaddress, self.scduser, self.scdpasswd, self.scdmount)
        self.lpcd = RemoteFolder(self.lpcdaddress, self.lpcduser, self.lpcdpasswd, self.lpcdmount)
        self.pcd = RemoteFolder(self.pcdaddress, self.pcduser, self.pcdpasswd, self.pcdmount)
        self.local = RemoteFolder(self.localaddress, self.localuser, self.localpasswd, self.localmount)

    def mount(self):
        for target in (self.rcd, self.scd, self.lpcd, self.pcd):
            netuse_mount = "net use {letter}: {IP} /user:{user} {passwd}".format(
                IP=target.address,
                user=target.user,
                passwd=target.password,
                letter=target.letter,
            )
            print(netuse_mount)

    def unmount(self):
        for target in (self.rcd, self.scd, self.lpcd, self.pcd):
            netuse_del = "net use {letter}: /delete".format(
                letter=target.letter,
            )
            print(netuse_del)

    def get_imagepath(self, rcd, scd, lpcd=None, patchcd=None):
        try:
            isolist = dict()
            softloadpn = re.compile(r'_(\w{2}\.\w{5}\.\w{3})_|\[(\w{2}\.\w{5}\.\w{3})\]', re.I)
            slircdpn = re.compile(r'(GRCD\d{14})', re.I)
            # for type, target in [("RCD", self.rcdmount), ("SCD", self.scdmount), ("LPCD", self.lpcdmount), ("PCD", self.pcdmount)]:
            for type, target in [("RCD", "F:\\SLIRCD"), ("SCD", "E:\\1_Image"), ("LPCD", "F:\\RCD\\20H1"),
                                 ("PCD", "D:\\1_PatchCD\\1_Released")]:
                isolist[type] = defaultdict(lambda: list())
                # folder = "{}:\\".format(target.letter)
                folder = target
                for root, dirs, files in os.walk(folder):
                    for file in files:
                        if file[-3:].lower() in ("iso", "swm"):
                            checkpn = softloadpn.search(file) or slircdpn.search(file)
                            if checkpn:
                                pn = checkpn.group(1) or checkpn.group(2)
                                file_path = os.path.join(root, file)
                                file_path = os.path.abspath(file_path)
                                isolist[type][pn].append(file_path)

            imageinfo = dict()
            print(rcd)
            if lpcd:
                imageinfo['Type'] = "RCD"
                u_rcd = rcd
                imageinfo['LPCD'] = list()
                for LPCD_NO in lpcd.split("/"):
                    if isolist['LPCD'][LPCD_NO]:
                        imageinfo['LPCD'].append(isolist['LPCD'][LPCD_NO][0])
                    else:
                        imageinfo['LPCD'] = None
                        break
            else:
                imageinfo['Type'] = "SLIRCD"
                imageinfo['softload'] = rcd
                u_rcd = slircdpn.search(rcd).group(1)
                print("u_rcd", u_rcd)

            imageinfo['RCD'] = isolist['RCD'][u_rcd] or None
            imageinfo['SCD'] = isolist['SCD'][scd] or None
            if patchcd:
                imageinfo['PCD'] = list()
                for Patch_NO in patchcd.split("/"):
                    if isolist['PCD'][Patch_NO]:
                        imageinfo['PCD'].append(isolist['PCD'][Patch_NO][0])
                    else:
                        imageinfo['PCD'] = None
                        break

            error_check = ""
            for key in imageinfo.keys():
                if imageinfo[key] is None:
                    error_check += "{} file(s) are found. /n".format(key)
            assert len(error_check) == 0, error_check
            return imageinfo
        except AssertionError as e:
            return e


class NAPPHDD(SQLConnect):
    def __init__(self, imageinfo):
        self.data = imageinfo
        self.xml = datetime.datetime.now().strftime('D:\\NappDeploySetting_%Y_%m_%d_%H-%M-%S.xml')

    def __autocheck__(self):
        os.system(r"start /w {tool} /autocheck:{xml}".format(tool=napphdd_exe, xml=self.xml))

    def __deploy__(self):
        os.system(r"start /w {tool} /deploy:{xml}".format(tool=napphdd_exe, xml=self.xml))

    def create_nappxml(self):
        try:
            XMLTree = dict()
            OS1_list = [
                ('RCDnode', 'RcdList'),
                ('SCDnode', 'Scd'),
                ('LPCDnode', 'LpcdList'),
                ('RSLKitnode', 'RSLKitID'),
                ('PCDnode', 'PcdList'),
                ('BomZip', 'BomZip')
            ]
            root = ET.Element('xNappDeploySetting')
            for r_node in ['FGSN', 'NewBIOSVersion', 'NewBIOSExecutePath', 'Os1Image']:
                OS1 = ET.SubElement(root, r_node)
            for node, node_name in OS1_list:
                XMLTree[node] = ET.SubElement(OS1, node_name)
            for other_node in ['IsFactoryImageForNAPP11', 'IsFIBorOSLRCD']:
                ET.SubElement(OS1, other_node).text = 'false'
            if self.data['Type'] == 'SLIRCD':
                ET.SubElement(OS1, 'IsSLIRCD').text = 'true'
            elif self.data['Type'] == 'RCD':
                ET.SubElement(OS1, 'IsSLIRCD').text = 'false'

            XMLTree['SCDnode'].text = self.data["SCD"]
            for RCD_n in self.data["RCD"]:
                print(RCD_n)
                ET.SubElement(XMLTree['RCDnode'], 'string').text = RCD_n
            for SCD_n in self.data["SCD"]:
                print(SCD_n)
                XMLTree['SCDnode'].text = SCD_n
            if "LPCD" in self.data.keys():
                for LPCD_n in self.data['LPCD']:
                    print(LPCD_n)
                    ET.SubElement(XMLTree['LPCDnode'], 'string').text = LPCD_n
            if "softload" in self.data.keys():
                XMLTree['RSLKitnode'].text = self.data['softload']
                # ET.SubElement(XMLTree['RSLKitnode'], 'string').text=LPCD_info
            if "PCD" in self.data.keys():
                for Patch_n in self.data["PCD"]:
                    print(Patch_n)
                    ET.SubElement(XMLTree['PCDnode'], 'string').text = Patch_n
        finally:
            tree = ET.ElementTree(root)
            tree.write(self.xml, encoding='utf-8')

    def install(self):
        try:
            self.__autocheck__()
            self.__deploy__()
        finally:
            # Check Stauts
            pass


def main():
    try:
        system = MachineInfo()
        sql_conn = SQLConnect()
        # scdlist = sql_conn.getscds_bymodelname(system.modelname)
        scdlist = sql_conn.getscds_bymodelname("TMP214")
        lastimage = scdlist[0]
        goldimage = [i for i in scdlist if i[2] == 'g']
        print(lastimage)
        print(goldimage)
        """
        # Call GUI to check user choose
        # Golden / Last
        # W10 PR / W10 Home
        # ENEU / TC
        """

        slircdinfo = sql_conn.getslircd_byscd("FD.DHSA0.00W", "TC", "65")
        remotehdd = RemoteHDD()
        image_path_info = remotehdd.get_imagepath(rcd=slircdinfo, scd="FD.DHSA0.00W", patchcd="FM.DRVD0.0PT/FM.DRVD0.0PR")
        NAPPHDD(image_path_info).create_nappxml()
    finally:
        print('Close')

main()
