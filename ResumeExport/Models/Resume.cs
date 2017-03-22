using System;
using System.Collections.Generic;
using System.Web.Mvc;

namespace ResumeExport.Models
{
    //履歷資料
    public class Resume
    {
        public string Name { get; set; }
        public string Gender { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
        public string Mobile { get; set; }
        [AllowHtml]
        public string Description1 { get; set; }
        [AllowHtml]
        public string Description2 { get; set; }
        public List<History> JobHistory { get; set; }

        //範例內容
        public Resume()
        {
            Name = "采威國際";
            Gender = "男";
            Email = "sample@iscom.com.tw";
            Address = "台中市南屯區公益路二段51號21樓";
            Phone = "04-23265200";
            Mobile = "0910000000";
            Description1 = "<p>我出生在一個很平凡但很美滿的小家庭，父親是個公務員，在台電服務，母親是個家庭主婦，而弟弟和我都還在學校求學。父母用民主的方式管教我們，希望我們能夠<strong>獨立自主、主動學習</strong>，累積人生經驗，但他們會適時的給予鼓勵和建議，父母親只對我們要求兩件事，第一是保持健康，第二是著重課業。因為沒有健康的身體，就算有再多的才華、再大的抱負也無法發揮出來。又因為家境並不富裕，所以必須專心於課業上，學得一技之長，將來才能自立更生。</p><p>小學時代的我很活潑、很好動，在課業上表現平平，但課外表現不錯，除了擔任過班長等幹部外，還參加樂隊、糾察隊等，另外還曾獲選為校隊參加跳高比賽。</p><p>小學畢業後，進入了一所私立中學，因為校規嚴格，使原本好動的我變得較為文靜，不過在那裡我學會了許多應有的禮節與待人處世的道理。在國中時期的我，好像開了竅，代表全校接受縣政府的表揚，在國三畢業典禮上，更代表了全體畢業生上台領取畢業證書。</p><p>進附中後，每天都覺得很充實、很快樂。附中學生的特色是能K、能玩，所以我不斷地努力學習，希望能夠達到此目標。在課業方面，我都能保持在一定的水準，因為上課專心聽講、仔細思考、體會老師所說的每一句話，在腦海裡架構重要觀念，一有問題就立刻發問，因此上課的吸收效率很高，不但使得複習的工作能夠很快完成，還有多餘的時間從事課外活動。在這麼多的科目當中，我最喜歡的是數學、化學和生物，因為數學、化學能夠訓練我們組織與思考能力。而生物則和日常生活有密切的關係，且它為我們揭開人體的奧秘。</p><p>我在學校人際關係良好，因為無論何時，總是笑臉迎人，微笑已成為我個人的正字招牌。老師們常稱讚我是個品學兼優的好學生，常給我許多的鼓勵，而同學間的相處十分融洽，彼此之間很有默契，團結一心。</p><p>除了課業之外，其他方面我也沒有偏廢。在高一時加入學校管樂隊，吹奏低音號，每天升旗參加出勤，在這裡不但使我對管樂器有進一步的認識，還認識了許多朋友，因此在那個時候，參加社團已成為我放學後的一大休閒。而我最喜歡的運動是棒球，我常在電視上或球場上觀賞職棒比賽，欣賞球員的美姿，模仿其動作。我常利用課餘時間約同學一起出外打球，記得在高二時，班上組隊參加班際壘球賽，那時我擔任隊長，防守二壘，首先展開攻勢，激起球隊士氣，最後終以一分之差贏得了最後勝利。除了棒球之外，我也很喜歡藍球、排球等，因此使得原本瘦弱的身體，變得強壯許多。</p><p><img src=\"https://s.yimg.com/bt/api/res/1.2/zPQUUimgO8yTF6Fk9ByVPw--/YXBwaWQ9eW5ld3NfbGVnbztjaD00OTk7Y3I9MTtjdz03NDc7ZHg9MDtkeT0wO2ZpPXVsY3JvcDtoPTEyNztxPTc1O3c9MTkw/http://media.zenfs.com/zh-Hant-TW/homerun/nownews.com/514d816da52298329f328923c28cd7a0\" /><br />澎湖「鯨魚洞」</p><p>我從小就立志要當醫生，近年來當我接觸了有關醫學系的相關資訊，漸漸地了解到當個醫生並不是那麼簡單的事，需要對服務人群有興趣，及擅長人際溝通，且在求學的過程中也相當辛苦。但疾病加諸在病人身上的痛苦是無法言諭的，來自醫生的關懷與勉勵，能讓病人產生無比的信念，勇敢地向疾病宣戰，在病人痊癒時，看到病人及家屬喜形於色，那種喜悅，令我十分嚮往，而且醫生不僅僅是要免除病人的疾病和虛弱，也要能兼顧對人們的整體關懷，使病患的身心達到安寧的狀態，在這一方面，它讓我更確定了讀醫學系的志向。</p>";

            List<History> history = new List<History>();
            history.Add(new History { CompanyName = "采威教育E化", JobTitle = "工讀生", StartDT = Convert.ToDateTime("2014/7/1"), EndDT = Convert.ToDateTime("2014/8/31") });
            history.Add(new History { CompanyName = "采威教育E化", JobTitle = "一級工程師", StartDT = Convert.ToDateTime("2015/1/1"), EndDT = Convert.ToDateTime("2016/12/31") });
            history.Add(new History { CompanyName = "采威教育E化", JobTitle = "二級工程師", StartDT = Convert.ToDateTime("2017/1/1") });
            JobHistory = history;
        }
    }


    //工作經歷
    public class History
    {
        public string CompanyName { get; set; }
        public string JobTitle { get; set; }
        public DateTime? StartDT { get; set; }
        public DateTime? EndDT { get; set; }
    }
}