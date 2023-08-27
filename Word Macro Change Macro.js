/**
 * Word改宏的宏 Macro
 * 宏由 Administrator 录制，时间: 2023/05/19
 所有更改替换均在以下两句，""双引号之内
      obj.Replacement.Text = "";
      Selection.Find.Replacement.Text = "";
 */
function Word改宏的宏() {
    /**
     * 
    yourIdOCR = `姓名王静
    性别女民族羌
    出生1986年7月11日
    住址辽宁省锦州市凌河区文政
    南里56-14号
    公民身份号码210703198607112029`;
    */
    yourIdOCR = `姓名左金强
    性别男民族汉
    出生1992年2月10日
    住址江苏省通州市平潮镇云台
    山村十四组36号
    公民身份号码320683199202106030`;
    //xueLiOCR = "入学日期\n1111年22月33日\n毕（结）业日期\n4444年55月66日\n学校名称\n1234大学\n专业\n4321\n学制\n3年\n层次\n专科\n学历类别\n普通高等教育\n学习形式\n普通全日制\n毕（结）业";
    //你的申报专业
    yourAppMajor = '给排水';
    //#region 助理工程师
    //你的助理工程师专业
    yourAsstEngrMajor = '';
    //你的助理工程师资格名称
    yourAsstEngrQualName = '';
    //yourAsstEngrQualName = '助理工程师';
    //你的助理工程师授予时间
    yourAsstEngrAwdTime = '';
    //你的助理工程师审批机关-人力资源和社会保障局
    yourAsstEngrApvlAuth = '';
    //#endregion
    //是(1)否(其他数字)全日制
    isFullTime = 1;
    //#region
    if (isFullTime == 1) {
        yourBracketDegFull = '（全日制）';//（日制度）
        yourfullTime = '是';//是否全日制
    } else {
        yourBracketDegFull = '（非全日制）';//（日制度）
        yourfullTime = '否';//是否全日制
    }
    //#endregion
    //你的参加工作时间
    year = 2014
    month = '08'
    //#region
    yourWrkHrs = year + '.' + month;
    //你的工作年限
    yourWrkLife = 2023 - year;
    //你的现从事工作岗位及时间
    yourCurrentJobTime = 3 + year + '.' + month;
    //#endregion
    //你的最高学历
    //yourHighestDegree = '专科';
    yourHighestDegree = '本科';
    //yourHighestDegree = '研究生';
    //yourHighestDegree = '博士生';

    //学位:学士(1)硕士(2)博士(3)空(其他)
    whatAcademicDegree = 1;
    //#region
    if (whatAcademicDegree == 1) {
        yourDegree = '学士';//学士
        yourBracketDegree = '科（学士）';//科（学士）
    } else if (whatAcademicDegree == 2) {
        yourDegree = '硕士';//硕士
        yourBracketDegree = '科（硕士）';//科（硕士）
    } else if (whatAcademicDegree == 3) {
        yourDegree = '博士';//博士
        yourBracketDegree = '科（博士）';//科（博士）
    } else {
        yourDegree = '';//空
        yourBracketDegree = '科';//空
    }
    //#endregion
    //你的毕业院校
    yourGraduatedCollege = '南京工业大学';
    //你的专业
    yourProfessional = '水质科学与技术';
    //你的毕业时间
    yourGraduationTime = '2014.06';
    //你的入学时间
    yourAdmissionTime = '2010.09';

    // 随机生成变量1，为数字在10到30之间
    variable1 = Math.floor(Math.random() * (30 - 10 + 1)) + 10;
    // 随机生成变量2，小于变量1且大于8
    variable2 = Math.floor(Math.random() * (variable1 - 9)) + 9;
    // 计算变量3，等于变量2减7
    variable3 = variable2 - 7;
    // 你的填表日期
    yourFillingDate = " " + variable1;
    // 你的公示时间，等于“2023.07.” + 变量3 + “-” + “2023.07.” + 变量2
    yourFairShowTime = "2023.07." + variable3.toString().padStart(2, '0') + "-2023.07." + variable2.toString().padStart(2, '0');

    //
    /**
     * 
    //#region 学历备案表
    // 给定的字符串
    var inputString = xueLiOCR
    // 使用正则表达式提取信息
    var regex = /入学日期\n(.+?)\n毕（结）业日期\n(.+?)\n学校名称\n(.+?)\n专业\n(.+?)\n学制\n(.+?)\n层次\n(.+?)\n学历类别\n(.+?)\n学习形式\n(.+?)\n毕（结）业/;
    var matches = inputString.match(regex);
 
    // 提取的信息
    var 入学日期 = matches[1];
    var 入学日期 = 入学日期.substr(0, 4) + "." + 入学日期.substr(5, 2);
 
    var 毕业日期 = matches[2];
    var 毕业日期 = 毕业日期.substr(0, 4) + "." + 毕业日期.substr(5, 2);
    
    //学校名称
    var 学校名称 = matches[3];
    //专业
    var 专业 = matches[4];
    //学制
    var 学制 = matches[5];
    //层次
    var 层次 = matches[6];
    //学历类别
    var 学历类别 = matches[7];
    //学习形式
    var 学习形式 = matches[8];
    //#endregion
    */
    //#region 身份证所有

    // 给定的字符串
    var str = yourIdOCR;

    // 提取姓名
    yourName = str.match(/姓名(.+)/)[1].trim();

    // 提取性别
    yourGender = str.match(/性别(.)(.+)/)[1];

    // 提取民族
    yourEthnicity = str.match(/民族(.+)/)[1].trim();

    // 提取号码后的内容
    //有无换行都行
    matchResult1 = str.match(/公民身份号码(.+)/); // 第一个匹配模式
    matchResult2 = str.match(/公民身份号码\n(.+)/); // 第二个匹配模式
    if (matchResult1 && matchResult1[1]) {
        yourIdNumber = matchResult1[1].trim();
    } else if (matchResult2 && matchResult2[1]) {
        yourIdNumber = matchResult2[1].trim();
    }

    matchResult = yourIdNumber.match(/^\d{6}(\d{4})(\d{2})/);
    if (matchResult) {
        birthYear = matchResult[1]; // 匹配结果的第一个捕获组是出生年份
        birthMonth = matchResult[2]; // 匹配结果的第二个捕获组是出生月份
    } else {
        // 提取出生日期
        yourBirthday = str.match(/出生(.+)/)[1].trim();
        // 提取年份
        birthYear = yourBirthday.match(/(\d{4})年/)[1];
        // 提取月份
        birthMonth = yourBirthday.match(/(\d{2})月/)[1];
    }
    yourBirthday = birthYear + "." + birthMonth
    //#endregion

    //你的居中名字
    if (yourName.length == 2) {
        yourCtrName = " " + yourName + " "; // 在变量name前后添加空格
    } else {
        yourCtrName = yourName;
    }
    //你的居中申报专业
    if (yourAppMajor.length == 2) {
        yourCtrAppMajor = "   " + yourAppMajor + "   "; // 在变量name前后添加空格
    } else if (yourAppMajor.length == 3) {
        yourCtrAppMajor = "  " + yourAppMajor + "  "; // 在变量name前后添加空格
    } else if (yourAppMajor.length == 4) {
        yourCtrAppMajor = " " + yourAppMajor + " "; // 在变量name前后添加空格
    } else {
        yourCtrAppMajor = yourAppMajor;
    }

    //#region 你的证明人
    //你的证明人1
    yourReference1 = '';
    //你的证明人2
    yourReference2 = '';
    // 人名列表
    var names = ["侯俊", "孙永", "苏镇", "唐康", "叶魏", "姚寒", "曹亮", "武合", "侯克", "谢玮", "叶船", "袁建", "段啸", "孙列", "邓国", "汤乒", "韩峰", "江益", "曾军", "余葆", "乔丞", "叶哲", "夏智", "丁致", "张卿", "杜锋", "彭强", "秦则", "刘剑", "乔宽", "唐辅", "汪荣", "罗竹", "郭茗", "孔茂", "邱娟", "田华", "胡霖", "蔡晶", "陆姗", "侯源", "卢峰", "董思", "郝群", "孙昱", "汪信", "汪民", "杨翰", "贾璟"];

    // 随机生成两个不重复的索引
    var nameIndex1 = Math.floor(Math.random() * names.length);
    var nameIndex2 = Math.floor(Math.random() * (names.length - 1));
    nameIndex2 = (nameIndex2 >= nameIndex1) ? nameIndex2 + 1 : nameIndex2; // 确保第二个索引不与第一个索引相同

    // 获取对应的人名
    yourReference1 = names[nameIndex1];
    yourReference2 = names[nameIndex2];
    //#endregion

    //#region 你的培训起止时间,你的培训地点
    //你的培训起止时间1
    yourTrngStrtstpTime1 = '';
    //你的培训起止时间2
    yourTrngStrtstpTime2 = '';
    //你的培训地点1
    yourTrngSite1 = '';
    //你的培训地点2
    yourTrngSite2 = '';
    // 你的培训起止时间
    var stringArray1 = ["2021.01-2021.04", "2021.02-2021.05", "2021.03-2021.06", "2021.04-2021.07", "2021.05-2021.08", "2021.06-2021.09", "2021.07-2021.10", "2021.08-2021.11", "2021.09-2021.12", "2021.10-2022.01", "2021.11-2022.02", "2021.12-2022.03", "2022.01-2022.04", "2022.02-2022.05", "2022.03-2022.06", "2022.04-2022.07", "2022.05-2022.08", "2022.06-2022.09", "2022.07-2022.10", "2022.08-2022.11", "2022.09-2022.12"]
    var randomElements1 = getRandomElements3(stringArray1, 2, 5);

    yourTrngStrtstpTime1 = randomElements1[0];
    yourTrngStrtstpTime2 = randomElements1[1];
    // 你的培训地点
    var stringArray2 = ["佳兆业中心", "诚大数码广场", "商会总部大厦", "中汇广场", "新华天玺大厦", "千缘财商汇", "城开中心", "OVU创客公社", "皇朝万鑫国际大厦", "欧亚联营商务大厦", "沈阳金谷科技园", "沈阳国际软件园", "嘉里中心企业广场", "美联大厦", "博宇金属大厦", "新世界丰盛商务大厦", "第一商城", "金廊万科中心", "沈阳信悦汇", "中海广场", "银河国际大厦", "恒隆广场", "新地中心", "北约客置地广场", "北方国际传媒中心", "卓越大厦", "华润大厦", "东北传媒文化广场", "华府大厦", "万鑫国际大厦", "茂业中心", "财富中心", "锦峰大厦", "天润广场", "总统大厦", "中信信悦汇广场", "仁恒置地广场", "SK大厦", "闽商总部大厦", "新世界商业中心"];
    var randomElements2 = getRandomElements2(stringArray2, 2);

    yourTrngSite1 = randomElements2[0];
    yourTrngSite2 = randomElements2[1];
    //#endregion

    //#region 助工证
    /**
     * 
    //你的助理工程师专业
    //
    //环境监测与治理
    //建筑工程管理
     */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的助理工程师专业";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourAsstEngrMajor;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /**
     * 
    //你的助理工程师资格名称
    //
    //助理工程师
     */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的助理工程师资格名称";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourAsstEngrQualName;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /**
     * 
    //你的助理工程师授予时间
    //
    //1970/1971/1972/1973/1974/1975/1976/1977/1978/1979
    //1980/1981/1982/1983/1984/1985/1986/1987/1988/1989
    //1990/1991/1992/1993/1994/1995/1996/1997/1998/1999
    //2000/2001/2002/2003/2004/2005/2006/2007/2008/2009
    //2010/2011/2012/2013/2014/2015/2016/2017/2018/2019
    //.01/.02/.03/.04/.05/.06/.07/.08/.09/.10/.11/.12
     */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的助理工程师授予时间";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourAsstEngrAwdTime;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /**
     * 
    //你的助理工程师审批机关
    //
    //人力资源和社会保障局
     */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的助理工程师审批机关";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourAsstEngrApvlAuth;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    //#endregion 

    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的民族";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourEthnicity;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /**
     * 
    //填写
    //（非全日制）
    //（全日制）
     */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "（日制度）";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourBracketDegFull;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /**
     * 
    //填写
    //是
    //否
     */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "是否全日制";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourfullTime;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    //名字
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的名字";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourName;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    //三条线居中名字/ 丁琳 /
    //（两字格式，前后各加一个空格）/ 伍六 /
    //（三字格式，不变）/伍六七/
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的居中名字";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourCtrName;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";

    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的申报专业";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourAppMajor;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    *
    //你的居中申报专业/ 道路桥隧 /
    //前后各3空格/   电气   /
    //专业名长短不一，不要强求
    //更长的不用加空格
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的居中申报专业";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourCtrAppMajor;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";

    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的性别";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourGender;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";

    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的出生年月";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourBirthday;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";

    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的身份证号码";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourIdNumber;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    //填写
    //1990/1991/1992/1993/1994/1995/1996/1997/1998/1999
    //2000/2001/2002/2003/2004/2005/2006/2007/2008/2009
    //2010/2011/2012/2013/2014/2015/2016/2017/2018/2019
    //.01/.02/.03/.04/.05/.06/.07/.08/.09/.10/.11/.12
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的参加工作时间";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourWrkHrs;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";

    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的工作年限";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourWrkLife;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    //填写(参加工作的时间+3~5年那个，只输入时间)例如2006.10
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的现从事工作岗位及时间";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourCurrentJobTime;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    //填写
    //专科
    //本科
    //研究生
    //博士生
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的最高学历";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourHighestDegree;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    //#region 学位
    /*
    //填写
    //科
    //科（学士）
    //科（硕士）
    //科（博士）
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "科（学位）";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourBracketDegree;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    //填写
    //
    //学士
    //硕士
    //博士
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的学位";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourDegree;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    //#endregion
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的毕业院校";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourGraduatedCollege;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的专业";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourProfessional;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的毕业时间";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourGraduationTime;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的入学时间";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourAdmissionTime;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的培训起止时间1";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourTrngStrtstpTime1;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的培训起止时间2";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourTrngStrtstpTime2;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的培训地点1";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourTrngSite1;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的培训地点2";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourTrngSite2;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的证明人1";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourReference1;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的证明人2";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourReference2;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    //填表日期
    空格加日期
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的填表日期";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourFillingDate;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
    /*
    //公示时间
    2023.01.01-2023.01.08-2023.01.15-2023.01.22-2023.01.29-2023.02.05-2023.02.12-2023.02.19-2023.02.26-2023.03.05-2023.03.12-2023.03.19-2023.03.26-2023.04.02-2023.04.09-2023.04.16-2023.04.23-2023.04.30-2023.05.07-2023.05.14-2023.05.21-2023.05.28-2023.06.04-2023.06.11-2023.06.18-2023.06.25
    2023.01.02-2023.01.09-2023.01.16-2023.01.23-2023.01.30-2023.02.06-2023.02.13-2023.02.20-2023.02.27-2023.03.06-2023.03.13-2023.03.20-2023.03.27-2023.04.03-2023.04.10-2023.04.17-2023.04.24-2023.05.01-2023.05.08-2023.05.15-2023.05.22-2023.05.29-2023.06.05-2023.06.12-2023.06.19-2023.06.26
    2023.01.03-2023.01.10-2023.01.17-2023.01.24-2023.01.31-2023.02.07-2023.02.14-2023.02.21-2023.02.28-2023.03.07-2023.03.14-2023.03.21-2023.03.28-2023.04.04-2023.04.11-2023.04.18-2023.04.25-2023.05.02-2023.05.09-2023.05.16-2023.05.23-2023.05.30-2023.06.06-2023.06.13-2023.06.20-2023.06.27
    2023.01.04-2023.01.11-2023.01.18-2023.01.25-2023.02.01-2023.02.08-2023.02.15-2023.02.22-2023.03.01-2023.03.08-2023.03.15-2023.03.22-2023.03.29-2023.04.05-2023.04.12-2023.04.19-2023.04.26-2023.05.03-2023.05.10-2023.05.17-2023.05.24-2023.05.31-2023.06.07-2023.06.14-2023.06.21-2023.06.28
    2023.01.05-2023.01.12-2023.01.19-2023.01.26-2023.02.02-2023.02.09-2023.02.16-2023.02.23-2023.03.02-2023.03.09-2023.03.16-2023.03.23-2023.03.30-2023.04.06-2023.04.13-2023.04.20-2023.04.27-2023.05.04-2023.05.11-2023.05.18-2023.05.25-2023.06.01-2023.06.08-2023.06.15-2023.06.22-2023.06.29
    2023.01.06-2023.01.13-2023.01.20-2023.01.27-2023.02.03-2023.02.10-2023.02.17-2023.02.24-2023.03.03-2023.03.10-2023.03.17-2023.03.24-2023.03.31-2023.04.07-2023.04.14-2023.04.21-2023.04.28-2023.05.05-2023.05.12-2023.05.19-2023.05.26-2023.06.02-2023.06.09-2023.06.16-2023.06.23-2023.06.30
    2023.01.07-2023.01.14-2023.01.21-2023.01.28-2023.02.04-2023.02.11-2023.02.18-2023.02.25-2023.03.04-2023.03.11-2023.03.18-2023.03.25-2023.04.01-2023.04.08-2023.04.15-2023.04.22-2023.04.29-2023.05.06-2023.05.13-2023.05.20-2023.05.27-2023.06.03-2023.06.10-2023.06.17-2023.06.24
    */
    Selection.Find.Wrap = wdFindContinue;
    Selection.Find.Wrap = wdFindContinue;
    (obj => {
        obj.Text = "你的公示时间";
        obj.Forward = true;
        obj.Wrap = wdFindContinue;
        obj.MatchCase = false;
        obj.MatchByte = true;
        obj.MatchWildcards = false;
        obj.MatchWholeWord = false;
        obj.MatchFuzzy = false;
        obj.Replacement.Text = yourFairShowTime;
    })(Selection.Find);
    (obj => {
        obj.Style = "";
        obj.Highlight = wdUndefined;
        (obj => {
            obj.Style = "";
            obj.Highlight = wdUndefined;
        })(obj.Replacement);
    })(Selection.Find);
    Selection.Find.Execute(undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, undefined, wdReplaceAll, undefined, undefined, undefined, undefined);
    Selection.Find.Replacement.Text = "";
}

function getRandomElements2(array, count) {
    var randomElements = [];
    var shuffledArray = array.slice(); // 创建数组副本以避免修改原始数组

    // 从数组中随机选择元素
    for (var i = 0; i < count; i++) {
        var randomIndex = Math.floor(Math.random() * shuffledArray.length);
        var element = shuffledArray.splice(randomIndex, 1)[0];
        randomElements.push(element);
    }

    return randomElements;
}
function getRandomElements3(array, count, minDistance) {
    var randomElements = [];

    if (count >= array.length || minDistance >= array.length) {
        return randomElements; // 如果需要选择的元素数量大于等于数组长度，或者最小距离大于等于数组长度，则返回空数组
    }

    var shuffledArray = array.slice(); // 创建数组副本以避免修改原始数组

    // 从数组中随机选择元素
    for (var i = 0; i < count; i++) {
        var randomIndex = Math.floor(Math.random() * shuffledArray.length);
        var element = shuffledArray.splice(randomIndex, 1)[0];
        randomElements.push(element);
    }

    // 检查随机选取的两个元素之间的距离是否小于最小距离，或者第二个元素在第一个元素之前，如果是，则递归调用函数重新选择元素
    if (Math.abs(array.indexOf(randomElements[0]) - array.indexOf(randomElements[1])) < minDistance || array.indexOf(randomElements[0]) > array.indexOf(randomElements[1])) {
        return getRandomElements3(array, count, minDistance);
    }

    return randomElements;
}
