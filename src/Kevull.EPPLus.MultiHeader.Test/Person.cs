using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace Kevull.EPPLus.MultiHeader.Test
{
    internal enum Gender
    {
        NotSpecified,
        Male,
        Female
    }

    internal record Person(string Name, string Surname, DateTime BirthDate, Gender Gender, int? NumOfComputers, Uri? Profile, TimeOnly AlarmTime)
    {
        internal static List<Person> BuildPeopleList()
        {
            return BuildPeopleList(0, 2);
        }

        internal static List<Person> BuildPeopleList(int skip, int take)
        {
            var result = new List<Person>
            {
                new Person("Médiamass","Large", DateTime.Parse("2017/05/28"), Gender.Male,  null, null, new TimeOnly(8,0)),
                new Person("Aimée","Bateson", DateTime.Parse("1958/06/07"), Gender.Female, 2, new Uri("https://github.com/"), new TimeOnly(7, 0)),
                new Person("Michelle", "California", DateTime.Parse("1991/06/04"), Gender.Female, null, new Uri("http://bloomberg.com/nulla/nisl/nunc/nisl.json"), new TimeOnly(303766036809)),
                new Person("Rudy", "Leile", DateTime.Parse("1995/10/25"), Gender.Male, 1, new Uri("http://gmpg.org/posuere/cubilia/curae/donec/pharetra/magna.jsp"), new TimeOnly(300347184016)),
                new Person("Aretha", "Staning", DateTime.Parse("1981/10/21"), Gender.Female, 2, null, new TimeOnly(392801647717)),
                new Person("Nicky", "Mangon", DateTime.Parse("1985/06/16"), Gender.Female, null, null, new TimeOnly(354851221336)),
                new Person("Shay", "Baybutt", DateTime.Parse("1993/10/06"), Gender.Male, 5, new Uri("http://omniture.com/leo/rhoncus/sed/vestibulum/sit/amet.jpg"), new TimeOnly(400511024696)),
                new Person("Papageno", "Corteis", DateTime.Parse("1976/10/11"), Gender.Male, 1, new Uri("https://economist.com/sit/amet/erat.jpg"), new TimeOnly(342684803082)),
                new Person("Odey", "Thunderman", DateTime.Parse("1998/04/17"), Gender.Male, null, new Uri("http://sina.com.cn/integer.js"), new TimeOnly(415877314254)),
                new Person("Sergei", "Gilmartin", DateTime.Parse("1999/06/02"), Gender.Male, 3, new Uri("https://ebay.co.uk/integer/ac/leo.json"), new TimeOnly(292893590861)),
                new Person("Montague", "Davidowsky", DateTime.Parse("1991/11/09"), Gender.Male, 3, new Uri("http://apple.com/justo/sollicitudin/ut/suscipit/a/feugiat.json"), new TimeOnly(266325117888)),
                new Person("Lotty", "Tomley", DateTime.Parse("1977/10/16"), Gender.Female, 1, new Uri("http://netscape.com/ut/erat/id/mauris/vulputate/elementum.png"), new TimeOnly(228426175270)),
                new Person("Peyton", "Fursey", DateTime.Parse("2000/06/20"), Gender.Male, 4, new Uri("https://wiley.com/vestibulum/vestibulum/ante.html"), new TimeOnly(414991391784)),
                new Person("Maura", "Flucks", DateTime.Parse("1989/09/18"), Gender.Female, 2, new Uri("http://samsung.com/eros/elementum/pellentesque/quisque/porta/volutpat.png"), new TimeOnly(259912902061)),
                new Person("Sibby", "Schach", DateTime.Parse("1981/10/22"), Gender.Female, 4, null, new TimeOnly (264448357793)),
                new Person("Allianora", "Akester", DateTime.Parse("1985/03/01"), Gender.Female, null, new Uri("https://i2i.jp/orci/mauris.html"), new TimeOnly(305188273455)),
                new Person("Dela", "Cotterill", DateTime.Parse("1999/02/21"), Gender.Female, 3, new Uri("https://economist.com/ultrices/posuere/cubilia/curae/donec.jpg"), new TimeOnly(227235895254)),
                new Person("Leanor", "Kentish", DateTime.Parse("1972/11/04"), Gender.Female, 3, new Uri("http://biblegateway.com/donec/ut/dolor/morbi/vel/lectus.jpg"), new TimeOnly(303027059095)),
                new Person("Jerrome", "Musselwhite", DateTime.Parse("1974/08/22"), Gender.Male, 3, new Uri("https://icq.com/amet/lobortis.xml"), new TimeOnly(313663947684)),
                new Person("Mal", "Greasley", DateTime.Parse("1987/02/16"), Gender.Male, 5, new Uri("http://multiply.com/nec/molestie/sed.aspx"), new TimeOnly(253306584027)),
                new Person("Sunny", "Folini", DateTime.Parse("1995/05/07"), Gender.Male, 3, null, new TimeOnly (413887949936)),
                new Person("Abramo", "Domokos", DateTime.Parse("1980/07/11"), Gender.Male, 3, new Uri("http://xrea.com/ligula/sit/amet/eleifend.json"), new TimeOnly(228601901535)),
                new Person("Isidora", "Hinks", DateTime.Parse("1977/03/27"), Gender.Female, 4, new Uri("https://independent.co.uk/platea/dictumst/aliquam/augue/quam/sollicitudin/vitae.png"), new TimeOnly(409097737153)),
                new Person("Jilleen", "Bimrose", DateTime.Parse("1980/12/06"), Gender.Female, 1, new Uri("http://edublogs.org/bibendum/felis.jpg"), new TimeOnly(260333992244)),
                new Person("Ingrim", "Dalrymple", DateTime.Parse("1974/12/25"), Gender.Male, 4, new Uri("http://google.it/tincidunt/eu.aspx"), new TimeOnly(294365164539)),
                new Person("Saidee", "Jorcke", DateTime.Parse("1990/08/17"), Gender.Female, null, new Uri("https://oracle.com/sociis/natoque/penatibus/et/magnis/dis/parturient.jsp"), new TimeOnly(358922732933)),
                new Person("Curry", "Field", DateTime.Parse("1983/07/27"), Gender.Male, 4, new Uri("https://yelp.com/interdum/mauris.jpg"), new TimeOnly(229087568285)),
                new Person("Linc", "Hallahan", DateTime.Parse("1973/05/25"), Gender.Male, null, new Uri("https://rediff.com/pede/justo.png"), new TimeOnly(216785456001)),
                new Person("Delmor", "Giscken", DateTime.Parse("1998/05/23"), Gender.Male, 2, new Uri("http://sciencedaily.com/luctus/et/ultrices/posuere.png"), new TimeOnly(279220239043)),
                new Person("Jake", "Rodda", DateTime.Parse("1997/12/18"), Gender.Male, 1, null, new TimeOnly (417642911775)),
                new Person("Rici", "Birkenhead", DateTime.Parse("1987/12/27"), Gender.Female, 5, new Uri("https://hexun.com/diam/neque/vestibulum/eget/vulputate/ut/ultrices.jsp"), new TimeOnly(385756738674)),
                new Person("Tiebout", "Skittle", DateTime.Parse("1983/05/30"), Gender.Male, 3, new Uri("https://chron.com/interdum/venenatis/turpis/enim/blandit.html"), new TimeOnly(363411046986)),
                new Person("Adriaens", "Rocca", DateTime.Parse("1974/01/25"), Gender.Female, 2, null, new TimeOnly (364486498516)),
                new Person("Gaelan", "Marns", DateTime.Parse("1989/11/13"), Gender.Male, null, new Uri("http://va.gov/at/nibh/in.json"), new TimeOnly(265819354204)),
                new Person("Fabe", "Petrelli", DateTime.Parse("1979/03/26"), Gender.Male, 2, new Uri("https://freewebs.com/sollicitudin/vitae/consectetuer/eget/rutrum.png"), new TimeOnly(234145922762)),
                new Person("Neysa", "Beckett", DateTime.Parse("1989/01/17"), Gender.Female, null, null, new TimeOnly (240964093598)),
                new Person("Jemmy", "Fifield", DateTime.Parse("1977/06/16"), Gender.Female, 2, null, new TimeOnly (352211451226)),
                new Person("Perla", "Krolak", DateTime.Parse("1978/10/29"), Gender.Female, 3, null, new TimeOnly (251816824778)),
                new Person("Margette", "Showers", DateTime.Parse("1975/12/13"), Gender.Female, 2, null, new TimeOnly (336681375416)),
                new Person("Bliss", "Ballsdon", DateTime.Parse("1983/08/14"), Gender.Female, null, new Uri("https://stanford.edu/tempor.json"), new TimeOnly(324593874244)),
                new Person("Lexi", "Biles", DateTime.Parse("1990/11/19"), Gender.Female, 2, new Uri("http://wsj.com/ac/nibh/fusce.json"), new TimeOnly(308166138551)),
                new Person("Stinky", "Caesar", DateTime.Parse("1998/10/29"), Gender.Male, null, null, new TimeOnly (228653478703)),
                new Person("Pepi", "McGeorge", DateTime.Parse("1988/07/24"), Gender.Female, 1, new Uri("http://statcounter.com/diam/vitae/quam/suspendisse/potenti.json"), new TimeOnly(333404263537)),
                new Person("Winnifred", "McPhelimy", DateTime.Parse("1972/11/02"), Gender.Female, null, null, new TimeOnly (304518477860)),
                new Person("Claudette", "Munford", DateTime.Parse("1976/11/15"), Gender.Female, null, new Uri("https://nyu.edu/gravida/sem.json"), new TimeOnly(406887394814)),
                new Person("Holmes", "Carnihan", DateTime.Parse("1995/04/20"), Gender.Male, null, null, new TimeOnly (258728149500)),
                new Person("Dex", "Roome", DateTime.Parse("1980/06/04"), Gender.Male, 2, new Uri("https://ox.ac.uk/libero/nullam/sit/amet/turpis.aspx"), new TimeOnly(276352542448)),
                new Person("Timmie", "Laughlin", DateTime.Parse("1993/03/04"), Gender.Male, 2, new Uri("http://elpais.com/at/vulputate/vitae/nisl/aenean/lectus.js"), new TimeOnly(391027237695)),
                new Person("Jaimie", "Lowden", DateTime.Parse("1975/06/18"), Gender.Female, 1, null, new TimeOnly (414334762170)),
                new Person("Waldo", "Hinnerk", DateTime.Parse("1986/07/25"), Gender.Male, null, new Uri("http://sohu.com/consequat/ut/nulla/sed/accumsan.js"), new TimeOnly(326596113424)),
                new Person("Bryanty", "Hartshorn", DateTime.Parse("1988/12/06"), Gender.Male, null, new Uri("https://last.fm/integer/ac/neque/duis/bibendum/morbi/non.xml"), new TimeOnly(273616320330)),
                new Person("Ashley", "Luscombe", DateTime.Parse("1980/02/21"), Gender.Male, 3, new Uri("http://twitter.com/tempor/convallis.json"), new TimeOnly(337357068841)),
                new Person("Franny", "Dederich", DateTime.Parse("1993/01/12"), Gender.Male, 5, null, new TimeOnly (246127553404)),
                new Person("Gage", "Whichelow", DateTime.Parse("1990/05/27"), Gender.Male, 5, new Uri("https://cam.ac.uk/diam.js"), new TimeOnly(220256775321)),
                new Person("Shirline", "Peschka", DateTime.Parse("1976/10/04"), Gender.Female, null, new Uri("https://hud.gov/aliquam/non/mauris/morbi.png"), new TimeOnly(336355972380)),
                new Person("Oswald", "Grabham", DateTime.Parse("1989/10/23"), Gender.Male, null, null, new TimeOnly (406355239727)),
                new Person("Rhetta", "Owbridge", DateTime.Parse("1994/07/12"), Gender.Female, 1, null, new TimeOnly (270119527132)),
                new Person("Ynes", "Lowy", DateTime.Parse("1983/01/18"), Gender.Female, 4, null, new TimeOnly (233225673083)),
                new Person("Darell", "Stickford", DateTime.Parse("1990/08/15"), Gender.Female, 5, null, new TimeOnly (422423377054)),
                new Person("Madelin", "Grabban", DateTime.Parse("1996/08/23"), Gender.Female, null, new Uri("http://google.com/non/interdum/in/ante/vestibulum/ante.js"), new TimeOnly(295507009143)),
                new Person("Izaak", "Garnsey", DateTime.Parse("1986/02/18"), Gender.Male, 1, new Uri("http://cbslocal.com/aenean/fermentum/donec/ut/mauris/eget.json"), new TimeOnly(302545422348)),
                new Person("Christy", "Jordi", DateTime.Parse("1983/09/22"), Gender.Male, 2, null, new TimeOnly (406464283262)),
                new Person("Nichol", "Brough", DateTime.Parse("1976/11/24"), Gender.Female, 2, new Uri("http://google.com/ut/dolor/morbi/vel/lectus.json"), new TimeOnly(291855107961)),
                new Person("Nixie", "Odlin", DateTime.Parse("1984/12/13"), Gender.Female, 4, new Uri("https://blogs.com/erat/volutpat.js"), new TimeOnly(397966793887)),
                new Person("Olivero", "Janousek", DateTime.Parse("1976/09/21"), Gender.Male, 4, new Uri("http://ucoz.com/lacinia/nisi/venenatis/tristique/fusce.html"), new TimeOnly(258760751467)),
                new Person("Cal", "Thornebarrow", DateTime.Parse("1984/12/24"), Gender.Male, 1, new Uri("http://yolasite.com/lectus/pellentesque/eget/nunc/donec/quis.aspx"), new TimeOnly(279215569483)),
                new Person("Krispin", "Badsworth", DateTime.Parse("1971/02/23"), Gender.Male, null, null, new TimeOnly (364675772764)),
                new Person("Newton", "Duling", DateTime.Parse("1986/04/18"), Gender.Male, 1, new Uri("http://instagram.com/nibh/quisque.png"), new TimeOnly(408692335213)),
                new Person("Geoff", "MacClenan", DateTime.Parse("1973/06/07"), Gender.Male, 2, new Uri("https://moonfruit.com/volutpat/in/congue.json"), new TimeOnly(420675813481)),
                new Person("Ofilia", "Poad", DateTime.Parse("1981/10/29"), Gender.Female, 3, new Uri("http://wiley.com/duis/bibendum/felis/sed.jpg"), new TimeOnly(370337638088)),
                new Person("Gay", "Schustl", DateTime.Parse("1991/10/07"), Gender.Male, 2, new Uri("https://wp.com/dis.json"), new TimeOnly(295141591726)),
                new Person("Warren", "Petren", DateTime.Parse("1993/04/13"), Gender.Male, 5, new Uri("http://sciencedirect.com/accumsan/felis/ut/at/dolor/quis.jsp"), new TimeOnly(296584057951)),
                new Person("Leila", "Manktelow", DateTime.Parse("1980/09/28"), Gender.Female, 3, new Uri("http://timesonline.co.uk/ut/blandit.png"), new TimeOnly(274288200236)),
                new Person("Gustaf", "Oldford", DateTime.Parse("1977/03/12"), Gender.Male, 2, null, new TimeOnly (252165580722)),
                new Person("Aksel", "Weldon", DateTime.Parse("1978/04/27"), Gender.Male, 1, new Uri("https://sbwire.com/purus/eu.xml"), new TimeOnly(244189581403)),
                new Person("Isiahi", "Andries", DateTime.Parse("1978/07/14"), Gender.Male, null, new Uri("https://google.de/tortor/risus/dapibus.html"), new TimeOnly(299295102104)),
                new Person("Nanete", "Bruff", DateTime.Parse("1970/12/20"), Gender.Female, 2, null, new TimeOnly (259573390783)),
                new Person("Sutherlan", "Gozney", DateTime.Parse("1988/03/28"), Gender.Male, 1, new Uri("http://webeden.co.uk/ac/est/lacinia/nisi/venenatis.json"), new TimeOnly(338757654695)),
                new Person("Cosmo", "Frawley", DateTime.Parse("1997/08/01"), Gender.Male, 3, null, new TimeOnly (297281374549)),
                new Person("Danyelle", "Daunter", DateTime.Parse("1975/07/22"), Gender.Female, null, new Uri("https://wikia.com/hac/habitasse/platea.jsp"), new TimeOnly(298399283424)),
                new Person("Conrade", "Greaves", DateTime.Parse("1976/08/29"), Gender.Male, 4, new Uri("https://pen.io/eget/massa/tempor/convallis/nulla/neque/libero.jsp"), new TimeOnly(425549926402)),
                new Person("Marcos", "Boocock", DateTime.Parse("1979/11/01"), Gender.Male, 3, null, new TimeOnly (286245320190)),
                new Person("Rivi", "Baudouin", DateTime.Parse("1990/02/25"), Gender.Female, 1, new Uri("https://hugedomains.com/vehicula/consequat.aspx"), new TimeOnly(403490831255)),
                new Person("Pete", "Brocklebank", DateTime.Parse("1988/04/02"), Gender.Male, 1, new Uri("https://comsenz.com/convallis/tortor/risus.js"), new TimeOnly(331039002257)),
                new Person("Oswald", "Ivanilov", DateTime.Parse("1993/02/26"), Gender.Male, 5, new Uri("http://wordpress.org/cursus/urna/ut/tellus/nulla/ut/erat.jsp"), new TimeOnly(300458056001)),
                new Person("Patrick", "Silverlock", DateTime.Parse("1990/05/02"), Gender.Male, 2, null, new TimeOnly (352635276713)),
                new Person("Benito", "Selcraig", DateTime.Parse("1984/05/16"), Gender.Male, null, new Uri("http://com.com/integer/tincidunt/ante.xml"), new TimeOnly(402980352044)),
                new Person("Joline", "Tincey", DateTime.Parse("1991/07/28"), Gender.Female, null, new Uri("http://mayoclinic.com/platea/dictumst/maecenas/ut/massa/quis.jsp"), new TimeOnly(388753703573)),
                new Person("Shari", "Boutwell", DateTime.Parse("1989/05/27"), Gender.Female, null, new Uri("https://theglobeandmail.com/aenean/sit/amet/justo/morbi.json"), new TimeOnly(327837702706)),
                new Person("Tedmund", "Sarfas", DateTime.Parse("1970/09/19"), Gender.Male, 1, new Uri("http://engadget.com/nunc/proin/at.xml"), new TimeOnly(230767920071)),
                new Person("Dewitt", "Lucien", DateTime.Parse("1993/02/27"), Gender.Male, 1, new Uri("http://technorati.com/habitasse/platea/dictumst/maecenas/ut.js"), new TimeOnly(386044451037)),
                new Person("Zita", "Soughton", DateTime.Parse("1991/10/19"), Gender.Female, 3, new Uri("https://edublogs.org/rhoncus/mauris/enim.jsp"), new TimeOnly(389163200446)),
                new Person("Tanitansy", "Bunclark", DateTime.Parse("1988/07/08"), Gender.Female, null, new Uri("http://mail.ru/consequat/metus/sapien/ut.jsp"), new TimeOnly(274200874109)),
                new Person("Kristen", "Pietrasik", DateTime.Parse("1994/10/12"), Gender.Female, null, new Uri("https://seattletimes.com/lorem/ipsum/dolor/sit/amet.jsp"), new TimeOnly(427898931186))
            };
            return result.Skip(skip).Take(take).ToList();
        }
    }


}
