const express = require('express');
var tempfile = require('tempfile');
var officegen = require('officegen');


app = express();


/* Claim */

app.get('/pretenziya', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=pretenziya.docx'
    });
    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = req.query[15], var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];
    
    const t1 = var1;
    var p1 = docx.createP({align: 'right'});
    p1.addText(t1, {font_size: '12', font_face: 'Times New Roman', bold: true});
    
    const t2 = var2;
    var p2 = docx.createP({align: 'right'});
    p2.addText(t2, {font_size: '12', font_face: 'Times New Roman', bold: true});

    const t3 = var3;
    var p3 = docx.createP({align: 'right'});
    p3.addText(t3, {font_size: '12', font_face: 'Times New Roman', bold: true});

    const t4 = var4;
    var p4 = docx.createP({align : 'right'});
    p4.addText(t4, {font_size: '12', font_face: 'Times New Roman', bold: true});

    var empty = docx.createP();

    const t5 = var5;
    var p5 = docx.createP({align: 'right'});
    p5.addText(t5, {font_size: '12', font_face: 'Times New Roman', bold: true});

    const t6 = var6;
    var p6 = docx.createP({align: 'right'});
    p6.addText(t6, {font_size: '12', font_face: 'Times New Roman', bold: true});

    const t7 = var7;
    var p7 = docx.createP({align : 'right'});
    p7.addText(t7, {font_size: '12', font_face: 'Times New Roman', bold: true});

    var empty = docx.createP();

    const const_text1 = "Представитель по доверенности:";
    var const_text1P = docx.createP({align : 'right'});
    const_text1P.addText(const_text1, {font_size : '12', font_face: 'Times New Roman', bold: true});

    
    const const_text2 = "Калдыбаев Чингисхан Каирбекович";
    var const_text2P = docx.createP({align : 'right'});
    const_text2P.addText(const_text2, {font_size : '12', font_face: 'Times New Roman', bold: true});

    
    const const_text3 = "ИИН 930727350230";
    var const_text3P = docx.createP({align : 'right'});
    const_text3P.addText(const_text3, {font_size : '12', font_face: 'Times New Roman', bold: true});

    
    const const_text4 = "Адрес: г. Астана, Мангилик Ел С 4.6  ";
    var const_text4P = docx.createP({align : 'right'});
    const_text4P.addText(const_text4, {font_size : '12', font_face: 'Times New Roman', bold: true});

    
    const const_text5 = "Сот тел: +77052819797";
    var const_text5P = docx.createP({align : 'right'});
    const_text5P.addText(const_text5, {font_size : '12', font_face: 'Times New Roman', bold: true});

    var emptyP = docx.createP();

    const t8 = "ПРЕТЕНЗИЯ";
    var p8 = docx.createP({align:'center'});
    p8.addText(t8, {font_size: '12', font_face: 'Times New Roman'});

    const t9 = "о расторжении договора";
    var p9 = docx.createP({align: 'center'});
    p9.addText(t9, {font_size : '12', font_face: 'Times New Roman'});

    const t10 = var8 + " между нами был заключен договор №"+var9+" о предоставлении займа, по условиям которого: общая сумма и валюта займа – "+var10+", срок займа – "+var11+", сумма возврата – "+var12+", включая вознаграждение "+var13+", стоимость услуги – "+var14+". Фактически погашенная сумма – тенге.";
    var p10 = docx.createP();
    p10.addText(t10, {font_size : '12', font_face: 'Times New Roman'});

    const t11 = "Считаю указанный договор незаконным, не соответствующим нормам действующего законодательства РК, нарушающим мои права и законные интересы,  и подлежащим признанию недействительным.";
    var p11 = docx.createP();
    p11.addText(t11, {font_size : '12', font_face: 'Times New Roman'});

    const t12 = "В соответствии с п.1 ст. 725-1 ГК РК, при нарушении обязательств по своевременному погашению займа, все платежи заемщика по договору займа, включая сумму вознаграждения, неустойки (штрафа, пени), комиссии и иных платежей, предусмотренных договором займа, за исключением предмете займа, в совокупности не могут превышать сумму выданного займа за весь период действия договора займа.";
    var p12 = docx.createP();
    p12.addText(t12, {font_size : '12', font_face: 'Times New Roman'});

    const t13 = "Обслуживание займа самостоятельным видом банковской услуги считаться не должно и не может, поскольку нет такого самостоятельного вида банковских операции, как обслуживание банковских займов, и она не удовлетворяет каких-либо потребностей клиента Банка. Наличие комиссии по обслуживание банковских займов по существу противоречит Закону РК «О защите прав потребителей».";
    var p13 = docx.createP();
    p13.addText(t13, {font_size : '12', font_face: 'Times New Roman'});

    const t14 = "Во-первых, в соответствии с пп.6 ст. 1 Закону РК «О защите прав потребителей» услуга - это деятельность, направленная на удовлетворение потребностей потребителей, результаты которой не имеют материального выражения. Во-вторых, исполнитель услуги не должен включать в договор с потребителем условия, которые нарушают и (или) ущемляют права потребителя. В данном случае, Банки в договорах банковских займов по существу не прописываются за какое именное банковское обслуживание осуществляет Банк по обслуживанию займа взимается комиссия, тогда как по закону о защите прав потребителя потребитель должен быть просвещен о получаемой услуге. Таким образом, Банки второго в действительности взимание комиссии производятся по непонятной простому человеку услуге.";
    var p14 = docx.createP();
    p14.addText(t14, {font_size : '12', font_face: 'Times New Roman'});

    const t15 = "Следует отметить, то, что Национальным Банком РК № 667/206/740 от 09-02-2012 г. было разъяснено, о необходимости прекращения практики взимания комиссии за ведение ссудного счета по заемным операциям со ссылкой на ЗРК «О платежах и переводах». В связи с этим, считаю комиссию за организацию и обслуживание займа незаконной.";
    var p15 = docx.createP();
    p15.addText(t15, {font_size : '12', font_face: 'Times New Roman'});

    const t16 = "Согласно п.1 ст.728 ГК РК при заключении договора банковского займа в качестве заимодателя выступает банк или иное юридическое лицо, имеющее лицензию уполномоченного государственного органа на предоставление займов в денежной форме.В соответствии со ст.4 Закона РК «О государственном регулировании, контроле и надзоре финансового рынка и финансовых организаций» не допускается осуществление профессиональной деятельности на финансовом рынке лицам, не имеющим соответствующей лицензии, выданной в соответствии с законодательством Республики Казахстан.";
    var p16 = docx.createP();
    p16.addText(t16, {font_size : '12', font_face: 'Times New Roman'});

    const t17 = "В силу п.3 ст.715 ГК РК юридическим лицам и гражданам запрещается привлечение денег в виде займа от граждан в качестве предпринимательской деятельности, и такие договоры признаются недействительными с момента их заключения.";
    var p17 = docx.createP();
    p17.addText(t17, {font_size : '12', font_face: 'Times New Roman'});

    const t18 = "В силу п.1 ст.159 ГК ничтожна сделка, совершенная без получения необходимого разрешения.";
    var p18 = docx.createP();
    p18.addText(t18, {font_size : '12', font_face: 'Times New Roman'});

    const t19 = "Из п.2 и 3 ст.157-1 ГК следует, что недействительная сделка не влечет юридических последствий, за исключением тех, которые связаны с ее недействительностью.";
    var p19 = docx.createP();
    p19.addText(t19, {font_size : '12', font_face: 'Times New Roman'});

    const t20 = "При недействительности сделки каждая из сторон обязана возвратить другой все полученное по сделке.";
    var p20 = docx.createP();
    p20.addText(t20, {font_size : '12', font_face: 'Times New Roman'});

    const t21 = "В связи с существенным нарушением условий (изменением обстоятельств) договора №"+var9+" от "+var8+" дальнейшее исполнение договора невозможно.";
    var p21 = docx.createP();
    p21.addText(t21, {font_size : '12', font_face: 'Times New Roman'});

    const t22 = "В случае несогласия с данной претензией, ждем от Вас письменный мотивированный ответ "+var15+". Ответ просим предоставить посредством электронной почты. В случае отсутствия ответа с Вашей стороны мы будем вынуждены обратиться в суд для решения данного вопроса в судебном порядке.";
    var p22 = docx.createP();
    p22.addText(t22, {font_size : '12', font_face: 'Times New Roman'});

    var empty = docx.createP();

    const t23 = "На основании изложенного, предлагаю:";
    var p23 = docx.createP();
    p23.addText(t23, {font_size : '12', font_face: 'Times New Roman'});

    const t24 = "     1. Договор №"+var9+" от "+var8+" о предоставлении займа расторгнуть."; 
    var p24 = docx.createP();
    p24.addText(t24, {font_size : '12', font_face: 'Times New Roman'});

    var empty = docx.createP();

    const t25 = "Приложение:";
    var p25 = docx.createP();
    p25.addText(t25, {font_size : '12', font_face: 'Times New Roman'});

    const t26 = "     1. Копия доверенности на имя Калдыбаева Ч.К.";
    var p26 = docx.createP();
    p26.addText(t26, {font_size : '12', font_face: 'Times New Roman'});

    const t27 = "     2. Копия удостоверения личности заявителя "+var5+".";
    var p27 = docx.createP();
    p27.addText(t27, {font_size : '12', font_face: 'Times New Roman'});

    const t28 = "Дата: "+var16+".";
    var p28 = docx.createP();
    p28.addText(t28, {font_size : '12', font_face: 'Times New Roman'});



    docx.generate(res);
});

/* ISK */

app.get('/docx', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=ISK.docx'
    });
    console.log(req.query);
    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = "(15)", var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];
    console.log(var4);  
    const param1 = var1;
    var courtP = docx.createP({align: 'right'});
    courtP.addText(param1, {font_size: '12', font_face: 'Times New Roman', fonts_spacing : '1'});
    
    const param2 = var2;
    var locationP = docx.createP({align: 'right'});
    locationP.addText(param2, {font_size: '12', font_face: 'Times New Roman'});
    
    var emptyP = docx.createP();
    
    const param3 = var3;
    var credintailsP = docx.createP({align : 'right'}); 
    credintailsP.addText(param3, {font_size: '12', font_face: 'Times New Roman'});

    const param4 = var4;
    var IINP  = docx.createP({align: 'right'});
    IINP.addText(param4, {font_size: '12', font_face: 'Times New Roman'});

    const param5 = var5;
    var c_locationP = docx.createP({align : 'right'});
    c_locationP.addText(param5, {font_size : '12', font_face: 'Times New Roman'});

    var emptyP = docx.createP();

    const const_text1 = "Представитель по доверенности:";
    var const_text1P = docx.createP({align : 'right'});
    const_text1P.addText(const_text1, {font_size : '12', font_face: 'Times New Roman'});

    
    const const_text2 = "Калдыбаев Чингисхан Каирбекович";
    var const_text2P = docx.createP({align : 'right'});
    const_text2P.addText(const_text2, {font_size : '12', font_face: 'Times New Roman'});

    
    const const_text3 = "ИИН 930727350230";
    var const_text3P = docx.createP({align : 'right'});
    const_text3P.addText(const_text3, {font_size : '12', font_face: 'Times New Roman'});

    
    const const_text4 = "Адрес: г. Астана, Мангилик Ел С 4.6  ";
    var const_text4P = docx.createP({align : 'right'});
    const_text4P.addText(const_text4, {font_size : '12', font_face: 'Times New Roman'});

    
    const const_text5 = "Сот тел: +77052819797";
    var const_text5P = docx.createP({align : 'right'});
    const_text5P.addText(const_text5, {font_size : '12', font_face: 'Times New Roman'});

    var emptyP = docx.createP();

    const param7 = var6;
    var dummyP7 = docx.createP({align : 'right'});
    dummyP7.addText(param7, {font_size : '12', font_face: 'Times New Roman'});

    
    const param8 = var7;
    var dummyTop1 = docx.createP({align : 'right'});
    dummyTop1.addText(param8, {font_size : '12', font_face: 'Times New Roman'});
    

    const param9 = var8;
    var dummyTop2 = docx.createP({align : 'right'});
    dummyTop2.addText(param9, {font_size : '12', font_face: 'Times New Roman'});

    var emptyP = docx.createP();

    const dummyP8 = docx.createP({align: 'center'});
    dummyP8.addText("ИСКОВОЕ   ЗАЯВЛЕНИЕ", {font_size : '12', bold : true, font_face: 'Times New Roman'});

    const title = "о признании оферты на предоставление займа (договор займа №" +var9+" от "+var10+" - недействительной, о приведении сторон в первоначальное положение.)";
    var dummyP9 = docx.createP({align : 'center'});
    dummyP9.addText(title, {font_size: '12', bold : true, font_face: 'Times New Roman'});

    const p1 = "      31 июля 2018 года офертой на предоставление займа (договор займа №"+var9+" от "+var10+")  (далее по тексту-договор), ответчиком истцу был предоставлен заем на сумму "+var11+" сроком на "+var12+" дня. Сумма к возврату составляла "+var13+" тенге с учетом комиссии за организацию займа в размере "+var14+" тенге.";
    var dummyP10 = docx.createP();
    dummyP10.addText(p1, {font_size: '12', font_face: 'Times New Roman'});
    
    const p2 = "      Считаю указанный договор-незаконным, не соответствующим нормам действующего законодательства РК, нарушающие права и законные интересы Истца, и подлежащему признанию недействительным, по нижеследующим основаниям.";
    var dummyP11 = docx.createP();
    dummyP11.addText(p2, {font_size: '12', font_face: 'Times New Roman'});

    const p3 = "      Согласно нормам действующего законодательства РК, закрепленных в п. 1 ст. 158 Гражданского кодекса: «Сделка, содержание которой не соответствует требованиям законодательства, а также сделка, совершенная с целью, заведомо противоречащей основам правопорядка, является оспоримой и может быть признана судом недействительной, если настоящим Кодексом и иными законодательными актами Республики Казахстан не установлено иное.»";
    var dummyP12 = docx.createP();
    dummyP12.addText(p3, {font_size: '12', font_face: 'Times New Roman'});

    const p4 = "     В свою очередь, понятие займа, закреплено в статье 715 ГК РКП «по договору займа одна сторона (заимодатель) передает, а в случаях, предусмотренных настоящим Кодексом или договором, обязуется передать в собственность (хозяйственное ведение, оперативное управление) другой стороне (заемщику) деньги или вещи, определенные родовыми признаками, а заемщик обязуется своевременно возвратить заимодателю такую же сумму денег или равное количество вещей того же рода и качества.»";
    var dummyP13 = docx.createP();
    dummyP13.addText(p4, {font_size: '12', font_face: 'Times New Roman'});

    const p5 = "     Законодателем, в статье 718 предусмотрено вознаграждение по договору займа: «Если иное не предусмотрено законодательными актами Республики Казахстан или договором, за пользование предметом займа заемщик выплачивает вознаграждение заимодателю в размерах, определенных договором.»";
    var dummyP14 = docx.createP();
    dummyP14.addText(p5, {font_size: '12', font_face: 'Times New Roman'});

    const p6 = "     2. Защита прав заемщиков банков, организаций, осуществляющих отдельные виды банковских операций, микрофинансовых организаций и кредитных товариществ обеспечивается путем установления предельного размера годовой эффективной ставки вознаграждения, включающей вознаграждение, все виды комиссий и иные платежи, взимаемые заимодателем в связи с выдачей и обслуживанием займа, и рассчитываемой в порядке, определенном законодательством Республики Казахстан.";
    var dummyP15 = docx.createP();
    dummyP15.addText(p6, {font_size: '12', font_face: 'Times New Roman'});

    const p7 = "     Предельный размер годовой эффективной ставки вознаграждения определяется нормативным правовым актом Национального Банка Республики Казахстан.";
    var dummyP16 = docx.createP();
    dummyP16.addText(p7, {font_size: '12', font_face: 'Times New Roman'});

    const p8 = "     На сегодняшний день, с целью защиты прав заемщика, предельная процентная ставка Национального Банка РК, утверждена в размере 56% (пятьдесят шесть процентов).";
    var dummyP17 = docx.createP();
    dummyP17.addText(p8, {font_size: '12', font_face: 'Times New Roman'});

    const p9 = "     И так, согласно условия договора займа, в котором процентная ставка вознаграждения не указана, вознаграждение выражено в денежной сумме равной "+var14+" тенге. В свою очередь правила выдачи займов, в п 5.4 указана процентная ставка в день от 2.5 до 2 процентов, что в год составляет 600-730% годовых. Что вопиюще не соответствует нормам действующего законодательства.";
    var dummyP18 = docx.createP();
    dummyP18.addText(p9, {font_size: '12', font_face: 'Times New Roman'});

    const p10 = "     Также, в ст. 8-1 Закона Республики Казахстан «О защите прав потребителей» условия, нарушающие права потребителей при заключении договора, в п.4 указано, «4) установление требования по оплате потребителем несоразмерно большой суммы (свыше тридцати процентов стоимости товара, услуги, работы) в случае невыполнения им обязательств по договору».";
    var dummyP19 = docx.createP();
    dummyP19.addText(p10, {font_size: '12', font_face: 'Times New Roman'});

    const p11 = "     И так в моем случае, заем получен "+var10+", сумма займа "+var11+", вознаграждение за пользование займом в течении "+var12+" дня – "+var14+" тенге.";
    var dummyP20 = docx.createP();
    dummyP20.addText(p11, {font_size: '12', font_face: 'Times New Roman'});

    const p12 = "     Сумма займа - "+var11;
    var dummyP21 = docx.createP();
    dummyP21.addText(p12, {font_size : '12', font_face: 'Times New Roman'});

    const p13 = "     Основной долг - "+var11;
    var dummyP22 = docx.createP();
    dummyP22.addText(p13, {font_size: '12', font_face: 'Times New Roman'});

    const p14 = "     Проценты – "+var14;
    var dummyP23 = docx.createP();
    dummyP23.addText(p14, {font_size: '12', font_face: 'Times New Roman'});

    const p15 = "     Штраф за просрочку – "+var16;
    var dummyP24 = docx.createP();
    dummyP24.addText(p15, {font_size: '12', font_face: 'Times New Roman'});

    const p16 = "     Продолжительность займа "+var17+" дней";
    var dummyP25 = docx.createP();
    dummyP25.addText(p16, {font_size: '12', font_face: 'Times New Roman'});

    const p17 = "     Погашенная сумма - "+var18;
    var dummyP26 = docx.createP();
    dummyP26.addText(p17, {font_size: '12', font_face: 'Times New Roman'});

    const p18 = "     Согласно расчету пени с учетом официальной ставки рефинансирования национального банка Республики Казахстан, пеня (штраф) составили "+var19+" тенге, Штраф за просрочку начисленный ответчиком равен "+var16+" тенге, что в INF раз превышает законную неустойку.";
    var dummyP27 = docx.createP();
    dummyP27.addText(p18, {font_size: '12', font_face: 'Times New Roman'});

    const p19 = "     В свою очередь, Гражданским Кодексом установлено, что «3) Осуществление гражданских прав не должно нарушать прав и охраняемых законодательством интересов других субъектов права, не должно причинять ущерб окружающей среде.";
    var dummyP28 = docx.createP();
    dummyP28.addText(p19, {font_size: '12', font_face: 'Times New Roman'});

    const p20 = "     4) Граждане и юридические лица должны действовать при осуществлении принадлежащих им прав добросовестно, разумно и справедливо, соблюдая содержащиеся в законодательстве требования, нравственные принципы общества, а предприниматели – также правила деловой этики. Эта обязанность не может быть исключена или ограничено договором.";
    var dummyP29 = docx.createP();
    dummyP29.addText(p20, {font_size: '12', font_face: 'Times New Roman'});

    const p21 = "     «Добросовестность, разумность и справедливость действий участников гражданских правоотношений предполагаются».";
    var dummyP30 = docx.createP();
    dummyP30.addText(p21, {font_size: '12', font_face: 'Times New Roman'});

    const p22 = "     Указанные нормы и в целом нормы законодательства ответчиком нарушены, условия договора противоречат основам правопорядка, нарушают права и законные интересы истца.";
    var dummyP31 = docx.createP();
    dummyP31.addText(p22, {font_size: '12', font_face: 'Times New Roman'});

    const p23 = "     Хотелось бы отметить, что в Концепции правовой реформы политики РК на период 2010 по 2020 годы, утвержднной Указом Президента РК на период 2010 по 2020 годы, утвержденной Указом Президента РК от 24 августа 2009 года №858, отмечено о необходимости совершенствования института признания сделок недействительными. Выработка законодательства стимулов добровольного страхования сделок, последующее признания, которых незаконными, содержит риск изъятия предмета сделки у одной из сторон.Уточнение понятия сделок, их состава и последствий неисполнения сделок.";
    var dummyP32 = docx.createP();
    dummyP32.addText(p23, {font_size: '12', font_face: 'Times New Roman'});

    const p24 = "     При этом подчеркнуто, что определение для Казахстана амбициозной цели – вхождения в 2050 году в число 30-ти самых развитых государств мира предъявляет высокие требования к национальной правовой системе, которая должна эффективно обеспечивать проводимый курс страны на повышение качества жизни человека , общества и укрепления государственности.";
    var dummyP33 = docx.createP();
    dummyP33.addText(p24, {font_size: '12', font_face: 'Times New Roman'});

    const p25 = "     В Республики Казахстан соданы благоприятные условия для инвесторской деятельности, именно поэтому Уважаемый суд, считаю необходимым уделить особое внимание такого рода сделкам (договорам займа с размером неустойки в 29.9 раз превышающие установленную законом неустойку) так как указанная деятельность порождает глобальное социальное неравенство, что в конечном счете отразится на экономике страны в целом, ибо зачем производить товар, оказывать услуги, когда можно выдавать займ на таких вот условиях.";
    var dummyP34 = docx.createP();
    dummyP34.addText(p25, {font_size: '12', font_face: 'Times New Roman'});

    const p26 = "     Согласно п.1 ст 76 Конституции Республики Казахстан, Судебная власть осуществляется от имени Республики Казахстан и имеет своим назначением защиту прав, свобод и законных интересов граждан и организаций, обеспечение исполнения Конституции, законов, иных нормативных правовых актов, международных договоров Республики.";
    var dummyP35 = docx.createP();
    dummyP35.addText(p26, {font_size: '12', font_face: 'Times New Roman'});

    const p27 = "     На основании вышеизложенного, в соответствии с нормами действующего законодательства Республики Казахстан,";
    var dummyP36 = docx.createP();
    dummyP36.addText(p27, {font_size:'12', font_face: 'Times New Roman'});
    
    const p111 = "     Ходатайство:";
    var dummyP111 = docx.createP();
    dummyP111.addText(p111, {font_size:'12', font_face: 'Times New Roman'});

    const p222 = "Уважаемый суд, прошу истребовать у ответчика договор займа №360921  от 10 февраля 2018 года,  в связи с тем, что мне не представляется его получить у ответчика.";
    var dummyP222 = docx.createP();
    dummyP222.addText(p222, {font_size:'12', font_face: 'Times New Roman'});

    const p28 = "ПРОШУ СУД:";
    var dummyP37 = docx.createP({align: 'center'});
    dummyP37.addText(p28, {font_size: '12', font_face: 'Times New Roman'});

    const p29 = "     1. Признать оферту на предоставление займа (договор займа №"+var9+" от "+var10+") заключенный между "+var6+" и "+var3+" – недействительной.";
    var dummyP38 = docx.createP();
    dummyP38.addText(p29, {font_size: '12', font_face: 'Times New Roman'});

    const p30 = "     2. Привести стороны по договору в первоначальное положение.";
    var dummyP39 = docx.createP();
    dummyP39.addText(p30, {font_size: '12', font_face: 'Times New Roman'});

    const p31 = "Приложение:";
    var dummyP40 = docx.createP();
    dummyP40.addText(p31, {font_size:'12', font_face: 'Times New Roman'});

    const p32 = "     1. Копия искового заявления для Ответчика";
    var dummyP41 = docx.createP();
    dummyP41.addText(p32, {font_size: '12', font_face: 'Times New Roman'});

    const p33 = "     2. Квитанция по оплате государственной пошлины";
    var dummyP42 = docx.createP();
    dummyP42.addText(p33, {font_size: '12', font_face: 'Times New Roman'});

    const p34 = "     3. Договор банковского займа от "+var10;
    var dummyP43 = docx.createP();
    dummyP43.addText(p34, {font_size: '12', font_face: 'Times New Roman'});

    const p35 = "     4. Досудебная претензия";
    var dummyP44 = docx.createP();
    dummyP44.addText(p35, {font_size: '12', font_face: 'Times New Roman'});

    const p36 = "     5. Скриншот об отправке претензии";
    var dummyP45 = docx.createP();
    dummyP45.addText(p36, {font_size: '12', font_face: 'Times New Roman'});

    const p37 = "     6. Доверенность на представление интересов";
    var dummyP46 = docx.createP();
    dummyP46.addText(p37, {font_size: '12', font_face: 'Times New Roman'});

    const p38 = "     7. Диплом о юридическом образовании";
    var dummyP47 = docx.createP();
    dummyP47.addText(p38, {font_size: '12', font_face: 'Times New Roman'});

    const p39 = "     8. Копия удостоверения личности "+var3;
    var dummyP48 = docx.createP();
    dummyP48.addText(p39, {font_size:'12', font_face: 'Times New Roman'});

    const p40 = "Представитель по доверенности                                  "
    const p41 = "    Калдыбаев Ч.К.";
    var dummyP49 = docx.createP();
    dummyP49.addText(p40, {font_size:'12', bold:true, font_face: 'Times New Roman'});
    dummyP49.addText(p41, {font_size:'12', bold:true, align:'right', font_face: 'Times New Roman'});

    docx.generate(res);
});


/* Objection */

app.get('/vozrazheniye', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Ispolnitelskaya%20nadpis.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = "(15)", var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];

    console.log(req.query);
    const t1 = var1;
    var p1 = docx.createP({align: 'right'});
    p1.addText(t1, {font_size : '12', font_face: 'Times New Roman'});

    const t2 = var2;
    var p2 = docx.createP({align: 'right'});
    p2.addText(t2, {font_size : '12', font_face: 'Times New Roman'});

    var emptyP = docx.createP();

    const t3 = var3;
    var p3 = docx.createP({align: 'right'});
    p3.addText(t3, {font_size : '12', font_face: 'Times New Roman'});

    const t4 = var4;
    var p4 = docx.createP({align: 'right'});
    p4.addText(t4, {font_size : '12', font_face: 'Times New Roman'});

    const t5 = var5;
    var p5 = docx.createP({align: 'right'});
    p5.addText(t5, {font_size : '12', font_face: 'Times New Roman'});

    const title = "Возражение";
    var titleP = docx.createP({align: 'center'});
    titleP.addText(title, {font_size : '12', font_face: 'Times New Roman'});

    const t6 = "Вами "+var6+" вынесена исполнительная надпись в отношении гражданина "+var7+", ИИН "+var8+", "+var9+", место рождения "+var10+",зарегистрированный по адресу: "+var11+" в пользу "+var12+", "+var13+", зарегистрированный по адресу "+var14+", согласно ст. 92-8 Закона Республики Казахстан «О нотариате» исполнительная надпись должна быть отменена в течении 3-х дней с момента поступления возражения.";
    var p6 = docx.createP();
    p6.addText(t6, {font_size : '12', font_face: 'Times New Roman'});

    const t7 = "Приложение : 1) копия извещения";
    var p7 = docx.createP();
    p7.addText(t7, {font_size : '12', font_face: 'Times New Roman'});

    const t8 = "                        2) Копия удостоверения личности\n\n";
    var p8 = docx.createP();
    p8.addText(t8, {font_size : '12', font_face: 'Times New Roman'});

    const t9 = "Представитель заявителя                                                                      Калдыбаев Ч.К.";
    var p9 = docx.createP();
    p9.addText(t9, {font_size : '12', font_face: 'Times New Roman'});

    var datetime = new Date();
    const t10 = "(на основании доверенности)                                                                            " +datetime.getDate()+"."+(datetime.getMonth() + 1)+"."+datetime.getFullYear().toString();
    var p10 = docx.createP();
    p10.addText(t10, {font_size : '12', font_face: 'Times New Roman'});

    docx.generate(res);
});

/* Petition 1 */

app.get('/petition', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Hodataystvo%20Otlozheniya.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = "(15)", var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];

    const t1 = var1;
    var p1 = docx.createP({align: 'right'});
    p1.addText(t1, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t2 = var2;
    var p2 = docx.createP({align: 'right', lSpacing : 1});
    p2.addText(t2, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t3 = var3;
    var p3 = docx.createP({align: 'right', indentTop: 0});
    p3.addText(t3, {font_size : '12', font_face: 'Times New Roman', bold : true});

    var empty = docx.createP();

    const t4 = var4;
    var p4 = docx.createP({align: 'right'});
    p4.addText(t4, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t5 = var5;
    var p5 = docx.createP({align: 'right'});
    p5.addText(t5, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t11 = "Представитель по доверенности:";
    var p11 = docx.createP({align: 'right'});
    p11.addText(t11, {font_size : '12', font_face: 'Times New Roman', bold : true});
    
    const t12 = "Калдыбаев Чингисхан Каирбекович";
    var p12 = docx.createP({align: 'right'});
    p12.addText(t12, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t13 = "ИИН 930727350230";
    var p13 = docx.createP({align: 'right'});
    p13.addText(t13, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t14 = "Адрес: г. Астана, Мангилик Ел С 4.6";
    var p14 = docx.createP({align: 'right'});
    p14.addText(t14, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t15 = "Конт.тел.: +77052819797";
    var p15 = docx.createP({align: 'right'});
    p15.addText(t15, {font_size : '12', font_face: 'Times New Roman', bold : true});

    
    var empty = docx.createP();

    const title = "Ходатайство.";
    const titleP = docx.createP({align: 'center'});
    titleP.addText(title, {font_size : '12', font_face: 'Times New Roman'});

    const t7 = "    Уважаемый суд, в Вашем производстве находится гражданское дело по исковому заявлению "+var7+" кому принадлежит гражданское дело к "+var8+" о признании договора краткосрочного займа недействительным. Прошу Вас отложить судебное заседание назначенное на "+var9+" , ввиду того что ни истец, ни представитель истца не могут явиться на процесс по уважительной причине.";
    var p7 = docx.createP({align: 'left'});
    p7.addText(t7, {font_size : '12', font_face: 'Times New Roman'});
    var empty = docx.createP();
    var empty = docx.createP();

    const t9 = "Представитель заявителя                                                                      Калдыбаев Ч.К.";
    var p9 = docx.createP();
    p9.addText(t9, {font_size : '12', font_face: 'Times New Roman'});

    var datetime = new Date();
    const t10 = "Дата: "+datetime.getDate()+"."+(datetime.getMonth() + 1)+"."+datetime.getFullYear().toString();
    var p10 = docx.createP();
    p10.addText(t10, {font_size : '12', font_face: 'Times New Roman'});

    docx.generate(res);
});


/* Petition 2 */

app.get('/petition2', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Hodataystvo%20Otlozheniya.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = "(15)", var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];

    const t1 = var1;
    var p1 = docx.createP({align: 'right'});
    p1.addText(t1, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t2 = var2;
    var p2 = docx.createP({align: 'right', lSpacing : 1});
    p2.addText(t2, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t3 = var3;
    var p3 = docx.createP({align: 'right', indentTop: 0});
    p3.addText(t3, {font_size : '12', font_face: 'Times New Roman', bold : true});

    var empty = docx.createP();

    const t4 = var4;
    var p4 = docx.createP({align: 'right'});
    p4.addText(t4, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t5 = var5;
    var p5 = docx.createP({align: 'right'});
    p5.addText(t5, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t11 = "Представитель по доверенности:";
    var p11 = docx.createP({align: 'right'});
    p11.addText(t11, {font_size : '12', font_face: 'Times New Roman', bold : true});
    
    const t12 = "Калдыбаев Чингисхан Каирбекович";
    var p12 = docx.createP({align: 'right'});
    p12.addText(t12, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t13 = "ИИН 930727350230";
    var p13 = docx.createP({align: 'right'});
    p13.addText(t13, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t14 = "Адрес: г. Астана, Мангилик Ел С 4.6";
    var p14 = docx.createP({align: 'right'});
    p14.addText(t14, {font_size : '12', font_face: 'Times New Roman', bold : true});

    const t15 = "Конт.тел.: +77052819797";
    var p15 = docx.createP({align: 'right'});
    p15.addText(t15, {font_size : '12', font_face: 'Times New Roman', bold : true});

    
    var empty = docx.createP();

    const title = "Ходатайство.";
    const titleP = docx.createP({align: 'center'});
    titleP.addText(title, {font_size : '12', font_face: 'Times New Roman'});

    const t7 = "    Уважаемый суд, в Вашем производстве находится гражданское дело по исковому заявлению "+var4+" к "+var6+" о признании оферты на предоставлении займа недействительной, о приведении сторон в первоначальное положение.";
    var p7 = docx.createP();
    p7.addText(t7, {font_size : '12', font_face: 'Times New Roman'});

    const t21 = "   Истец и представитель истца поддерживают исковые требования в полном объеме. Ввиду невозможности участия в судебном заседании, просим рассмотреть дело без нашего участия. Копию судебного акта прошу направить по адресу представителя истца:\n\nг. Астана, Мангилик Ел С 4.6.";
    var p21 = docx.createP()
    p21.addText(t21, {font_size : '12', font_face: 'Times New Roman'});
    
    var empty = docx.createP();
    var empty = docx.createP();

    const t9 = "Представитель заявителя                                                                      Калдыбаев Ч.К.";
    var p9 = docx.createP();
    p9.addText(t9, {font_size : '12', font_face: 'Times New Roman'});

    var datetime = new Date();
    const t10 = "Дата: "+datetime.getDate()+"."+(datetime.getMonth() + 1)+"."+datetime.getFullYear().toString();
    var p10 = docx.createP();
    p10.addText(t10, {font_size : '12', font_face: 'Times New Roman'});

    docx.generate(res);
});


/* Судебный приказ  */


app.get('/debitorka', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Debitorka%20Sudebniy%20Prikaz.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = "(15)", var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];

    const title = "СУДЕБНЫЙ ПРИКАЗ";
    var titleP = docx.createP({align: 'center'});
    titleP.addText(title, {font_size : '14', font_face: 'Times New Roman', bold: true});

    const t2 = "28 декабря 2018 года          	№ "+var1+"	г."+var2;
    var p2 = docx.createP();
    p2.addText(t2, {font_size : '14', font_face: 'Times New Roman'});

    const t3 = "    Судья Специализированного межрайонного экономического суда "+var3+", рассмотрев заявление взыскателя "+var4+" БИН "+var5+" о вынесении судебного приказа о взыскании долга в сумме "+var6+" государственную пошлину в сумме 794 тенге с должника "+var7+" в лице "+var8+", ИИН "+var9+" должник адрес: "+var10+", ВП-38 на основании договора на оказание услуг по вывозу твердых бытовых отходов № "+var11+" на оказание услуг от "+var12+" года, руководствуясь статьями 139, 140 и 141 Гражданского процессуального кодекса Республики Казахстан (далее - ГПК), судья";
    var p3 = docx.createP();
    p3.addText(t3, {font_size : '14', font_face: 'Times New Roman'});

    const t4 = "ПРИКАЗЫВАЮ:";
    var p4 = docx.createP({align: 'center'});
    p4.addText(t4, {font_size : '14', font_face: 'Times New Roman', bold: true});

    const t5 = "    Взыскать с "+var7+" в лице "+var8+" в пользу  "+var4+" задолженность в сумме "+var6+" тенге, расходы по уплате государственной пошлины в сумме 794 (семьсот девяносто четыре) тенге.";
    var p5 = docx.createP();
    p5.addText(t5, {font_size : '14', font_face: 'Times New Roman'});

    const t6 = "    Разъяснить должнику, что в соответствии с частью 2 статьи 141 ГПК он вправе в течение десяти рабочих дней со дня, когда ему стало известно о вынесении судебного акта направить в суд вынесший судебный акт возражение против заявленного требования.";
    var p6 = docx.createP();
    p6.addText(t6, {font_size : '14', font_face: 'Times New Roman'});

    const t7 = "    Настоящий судебный приказ имеет силу исполнительного документа согласно части 2 статьи 134 ГПК.";
    var p7 = docx.createP();
    p7.addText(t7, {font_size : '14', font_face: 'Times New Roman'});

    const t8 = "Судья                                                                                	"+var3;
    var p8 = docx.createP();
    p8.addText(t8, {font_size : '14', font_face: 'Times New Roman'});

    const t9 = "Копия верна:";
    var p9 = docx.createP({align: 'left'});
    p9.addText(t9, {font_size : '14', font_face: 'Times New Roman'});

    const t10 = "Судья                                                                                	"+var3;
    var p10 = docx.createP();
    p10.addText(t10, {font_size : '14', font_face: 'Times New Roman'});

    docx.generate(res);
});

app.get('/courtorder', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Sudebniy%20Prikaz.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = req.query[15], var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];
    const right = ""+var1;
    var rightP = docx.createP({align: 'right'});
    rightP.addText(right, {font_size : '14', font_face: 'Times New Roman'})

    const title = "СУДЕБНЫЙ ПРИКАЗ";
    var titleP = docx.createP({align: 'center'});
    titleP.addText(title, {font_size : '14', font_face: 'Times New Roman', bold: true});

    const t1 = var2+"                                                             г.Астана";
    var p1 = docx.createP();
    p1.addText(t1 , {font_size : '14', font_face: 'Times New Roman'});
    var5 = var16+","+var15;
    var6 = var17+","+var18+","+var19;
    
    const t2 = "    Судья "+var3+" "+var4+", рассмотрев заявление "+var5+" о вынесении судебного приказа о взыскании с "+var6+" задолженности за эксплуатационные услуги, руководствуясь статьями 109, 134, подпунктом 10) статьи 135, 140 Гражданского процессуального кодекса Республики Казахстан (далее ‑ ГПК), статьей 272 Гражданского кодекса Республики Казахстан, статьей 50 Закона Республики Казахстан «О жилищных отношениях»";
    var p2 = docx.createP();
    p2.addText(t2, {font_size : '14', font_face: 'Times New Roman'});

    const t3 = "ПРИКАЗЫВАЮ:";
    var p3 = docx.createP({align: 'center'});
    p3.addText(t3, {font_size : '14', font_face: 'Times New Roman'  });

    const t4 = "Взыскать с "+var6+" в пользу "+var5+" задолженность "+var7+", расходы по уплате государственной пошлины 8.в сумма гос. пошлины , всего "+var9+" тенге.";
    var p4 = docx.createP();
    p4.addText(t4, {font_size : '14', font_face: 'Times New Roman'});

    const t5 = "Разъяснить должнику, что в соответствии с частью 2 статьи 141 ГПК он вправе в течение десяти рабочих дней со дня получения копии судебного приказа или со дня, когда ему стало известно о его вынесении, направить в суд вынесший судебный акт возражение против заявленного требования.";
    var p5 = docx.createP();
    p5.addText(t5, {font_size : '14', font_face: 'Times New Roman'});

    const t6 = "Судебный приказ имеет силу исполнительного листа.";
    var p6 = docx.createP();
    p6.addText(t6, {font_size : '14', font_face: 'Times New Roman'});
    console.log("asd");

    const t7 = "Судья                                           	"+var4;
    var p7 = docx.createP();
    p7.addText(t7, {font_size : '14', font_face: 'Times New Roman'});
    console.log("asd");

    const t8 = "Копия верна:";
    var p8 = docx.createP();
    p8.addText(t8, {font_size : '14', font_face: 'Times New Roman'});
    console.log("asd");

    const t9 = "Судья                                             	"+var4;
    var p9 = docx.createP();
    p9.addText(t9, {font_size : '14', font_face: 'Times New Roman'});
    console.log("asd");

    docx.generate(res);
});

///
app.get('/precourtclaim', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Dosudebnaya%20Pretenziya.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = req.query[15], var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];
    
    const t1 = ""+var1;
    var p1 = docx.createP({align: 'center'});
    p1.addText(t1, {font_size : '14', font_face: 'Times New Roman'});

    const t2 = ""+var2;
    var p2 = docx.createP({align: 'center'});
    p2.addText(t2, {font_size : '14', font_face: 'Times New Roman'});

    const t3 = ""+var3;
    var p3 = docx.createP({align: 'right'});
    p3.addText(t3, {font_size : '14', font_face: 'Times New Roman'});

    const t4 = ""+var4;
    var p4 = docx.createP({align: 'right'});
    p4.addText(t4, {font_size : '14', font_face: 'Times New Roman'});

    const t5 = ""+var5;
    var p5 = docx.createP({align: 'right'});
    p5.addText(t5, {font_size : '14', font_face: 'Times New Roman'});

    const t6 = "Досудебная претензия";
    var p6 = docx.createP({align: 'center'});
    p6.addText(t6, {font_size : '14', font_face: 'Times New Roman'});

    const t7 = "Уважаемая "+var3;
    var p7 = docx.createP({align: 'center'});
    p7.addText(t7, {font_size : '14', font_face: 'Times New Roman'});

    const t8 = var1+" просит вас погасить задолженность в размере "+var6+" (расходы на содержание общего имущества кондоминиума).";
    var p8 = docx.createP();
    p8.addText(t8, {font_size : '14', font_face: 'Times New Roman'});

    const t9 = "Напоминаем что согласно пункту 4 статьи 50 Закона РК «О Жилищных отношениях» при просрочке собственниками помещений обязательных платежей в счет общих расходов за каждый просроченный день, начиная с первого дня последующего месяца, на сумму долга начисляется пеня в размере установленным законодательством. При непогашении задолженности "+var1+" вправе обратиться в суд о принудительном взыскании задолженности.";
    var p9 = docx.createP();
    p9.addText(t9, {font_size : '14', font_face: 'Times New Roman'});

    const t10 = var1+" оставляет за собой право предоставить Вам в срок до "+var7+" на претензию возможность погасить имеющуюся задолженность, в противном случае данное уведомление будет считаться досудебным урегулированием спора и официальным уведомлением Вас о принятых в отношении Вас мер правового характера.";
    var p10 = docx.createP();
    p10.addText(t10, {font_size : '14', font_face: 'Times New Roman'});

    const t11 = "   Оплату можно произвести по нижеуказанным реквизитам.";
    var p11 = docx.createP();
    p11.addText(t11, {font_size : '14', font_face: 'Times New Roman'});
    const t18 = "   По всем вопросам вы можете обратиться "+var1+" "+var2;
    var p18 = docx.createP();
    p18.addText(t18, {font_size : '14', font_face: 'Times New Roman'});
    const t12 = "   Реквизиты:";
    var p12 = docx.createP();
    p12.addText(t12, {font_size : '14', font_face: 'Times New Roman'});
    
    const t13 = "   БИН:"+var8;
    var p13 = docx.createP();
    p13.addText(t13, {font_size : '14', font_face: 'Times New Roman'});

    const t14 = "   ИИК:"+var9;
    var p14 = docx.createP();
    p14.addText(t14, {font_size : '14', font_face: 'Times New Roman'});

    const t15 = "   БИК:"+var10;
    var p15 = docx.createP();
    p15.addText(t15, {font_size : '14', font_face: 'Times New Roman'});

    const t16 = "   "+var11;
    var p16 = docx.createP();
    p16.addText(t16, {font_size : '14', font_face: 'Times New Roman'});

    const t17 ="Представитель по доверенности                                  Калдыбаев Ч.К.";
    var p17 = docx.createP();
    p17.addText(t17, {font_size : '14', font_face: 'Times New Roman'});


    docx.generate(res);
});
///
////
app.get('/contract1', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Dosudebnaya%20Pretenziya.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = req.query[15], var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19];
    var datetime = new Date();
    let datestring = ""+datetime.getDate()+"."+(datetime.getMonth() + 1)+"."+datetime.getFullYear().toString();
    const title1 = "ДОГОВОР № " + var1;
    var titleP = docx.createP({align: 'center'});
    titleP.addText(title1, {bold: true, font_face: 'Times New Roman', font_size: '12'});
    
    const title2 = "на оказание юридических услуг";
    var title2P = docx.createP({align: 'center'});
    title2P.addText(title2 ,{bold: true, font_face: 'Times New Roman', font_size: '12'});
    
    const t1 = "г.Астана                                                                                         " + datestring;
    const p1 = docx.createP();
    p1.addText(t1,{font_face: 'Times New Roman', font_size: '12'});

    const t2 = "ТОО «GEN4», БИН: 150240008616, в лице Директора – Іскендір Нұрхан Қонысбайұлы, действующего на основании Устава, именуемое в дальнейшем «Исполнитель» с одной стороны и "+var3+" именуемый в дальнейшем «Заказчик» с другой стороны, заключили настоящий договор о нижеследующем.";
    var p2 = docx.createP();
    p2.addText(t2,{font_face: 'Times New Roman', font_size: '12'});

    const title3 = "1.Предмет договора";
    var title3P = docx.createP({align: 'center'});
    title3P.addText(title3,{bold:true,font_face: 'Times New Roman', font_size: '12'});

    const t3 = "1.1.Заказчик  поручает, а Исполнитель  принимает на себя обязательство по оказанию юридической помощи по признанию договора займа  между заказчиком и Микро финансовыми организациями недействительным в части.";
    var p3 = docx.createP();
    p3.addText(t3,{font_face: 'Times New Roman', font_size: '12'});

    const t4 = "1.2. Исполнитель  оказывает Заказчику по необходимости, следующую юридическую помощь: консультации, разъяснения, советы, собеседования, выезд, собирание фактических данных, ознакомление с предоставленными материалами и делами, предоставление письменных заключений, т.е. полностью занимается подготовкой, сбором всех документов необходимых для защиты прав Заказчика.";
    var p4 = docx.createP();
    p4.addText(t4,{font_face: 'Times New Roman', font_size: '12'});

    const t5 = "1.3. Исполнитель обязуется предпринять все законные меры, необходимые для достижения положительного результата, однако не дает конкретных гарантий на какой – либо определенный исход дела.";
    var p5 = docx.createP(t5);
    p5.addText(t5,{font_face: 'Times New Roman', font_size: '12'});

    const t6 = "1.4. Исполнитель может отступить от условий обязательства, пожеланий и требований Заказчика, если его действия совершены в интересах последнего, в рамках действующего законодательства, а также при отсутствии возможности незамедлительно связаться и согласовать с Заказчиком эти действия.";
    var p6 = docx.createP();
    p6.addText(t6,{font_face: 'Times New Roman', font_size: '12'});

    const t7 = "1.5. Исполнитель оказывает юридические услуги Заказчику по представительству, от имени последнего, в органах прокуратуры, Департаменте внутренних дел, Департаменте по исполнению судебных актов, судебных органах, налоговых и правоохранительных органах, а также, в государственных и негосударственных организациях, учреждениях, предприятиях и иных хозяйственных субъектах всех форм собственности, их структурных подразделениях по всем вопросам, связанных с защитой интересов Заказчика в Республике Казахстан.";
    var p7 = docx.createP();
    p7.addText(t7,{font_face: 'Times New Roman', font_size: '12'});

    const t8 = "1.6.Заказчик производит оплату в размере и сроки предусмотренные настоящим договором.";
    var p8 = docx.createP();
    p8.addText(t8,{font_face: 'Times New Roman', font_size: '12'});

    const title4 = "2. Права и обязанности Сторон";
    var title4P = docx.createP();
    title4P.addText(title4,{font_face: 'Times New Roman', bold: true, font_size: '12'});

    const t9 = "2.1. права и обязанности Исполнителя";
    var p9 = docx.createP();
    p9.addText(t9, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    const t10 = "2.1.1. Выполнить предоставленные услуги, предусмотренные предметом настоящего договора, надлежащим образом.";
    var p10 = docx.createP();
    p10.addText(t10,{font_face: 'Times New Roman', font_size: '12'});

    const t11 = "2.1.2. Вести дела Заказчика во всех судебных учреждениях, на различных стадиях судебного разбирательства, включая:";
    var p11 = docx.createP();
    p11.addText(t11,{font_face: 'Times New Roman', font_size: '12'});

    const t12 = "а)подготовку дел к судебному разбирательству;";
    var p12 = docx.createP();
    p12.addText(t12,{font_face: 'Times New Roman', font_size: '12'});

    const t13 = "б)заявлять ходатайства;";
    var p13 = docx.createP();
    p13.addText(t13,{font_face: 'Times New Roman', font_size: '12'});

    const t14 = "в)давать отзывы, возражения на заявления и ходатайства других лиц;";
    var p14 = docx.createP();
    p14.addText(t14,{font_face: 'Times New Roman', font_size: '12'});

    const t15 = "г)знакомиться с материалами дела;";
    var p15 = docx.createP();
    p15.addText(t15,{font_face: 'Times New Roman', font_size: '12'});

    const t16 = "д)подписывать и подавать исковые заявления;";
    var p16 = docx.createP();
    p16.addText(t16,{font_face: 'Times New Roman', font_size: '12'});

    const t17 = "е)истребовать документы, получать и давать ответы;";
    var p17 = docx.createP();
    p17.addText(t17,{font_face: 'Times New Roman', font_size: '12'});

    const t18 = "ж)подписывать, необходимые, процессуальные и иные документы;";
    var p18 = docx.createP();
    p18.addText(t18,{font_face: 'Times New Roman', font_size: '12'});

    const t19 = "з)подавать претензии, иски, встречные иски;";
    var p19 = docx.createP();
    p19.addText(t19,{font_face: 'Times New Roman', font_size: '12'});

    const t20 = "и)представлять Заказчика в гражданском и уголовном судопроизводстве, со всеми правами предоставленными законом истцу, ответчику, третьему лицу и потерпевшему;";
    var p20 = docx.createP();
    p20.addText(t20,{font_face: 'Times New Roman', font_size: '12'});

    const t21 = "к)обжаловать решения судов в апелляционном и надзорном порядке;";
    var p21 = docx.createP();
    p21.addText(t21,{font_face: 'Times New Roman', font_size: '12'});

    const t22 = "л)предъявлять исполнительный лист ко взысканию и представлять интересы Заказчика в исполнительном производстве;";
    var p22 = docx.createP();
    p22.addText(t22,{font_face: 'Times New Roman', font_size: '12'});

    const t23 = "м)расписываться за Заказчика и совершать все действия связанные с выполнением настоящего Договора;";
    var p23 = docx.createP();
    p23.addText(t23,{font_face: 'Times New Roman', font_size: '12'});

    const t24 = "н)занимать процессуальную позицию, не противоречащую интересам и не ухудшающую положение Заказчика;";
    var p24 = docx.createP();
    p24.addText(t24,{font_face: 'Times New Roman', font_size: '12'});

    const t25 = "о)соблюдать конфиденциальность.";
    var p25 = docx.createP();
    p25.addText(t25,{font_face: 'Times New Roman', font_size: '12'});

    const t26 = "2.1.3.Уведомлять Заказчика и их доверенных лиц о ходе судебного разбирательства, о его результатах, об  исполнении настоящего Договора не реже чем один раз в неделю.";
    var p26 = docx.createP();
    p26.addText(t26,{font_face: 'Times New Roman', font_size: '12'});

    const t27 = "2.1.4.Уведомлять Заказчика и их доверенных лиц о дне, времени и месте слушания гражданского дела.";
    var p27 = docx.createP();
    p27.addText(t27,{font_face: 'Times New Roman', font_size: '12'});

    const t28 = "2.2. Права и обязанности Заказчика";
    var p28 = docx.createP();
    p28.addText(t28, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    const t29 = "2.2.1. При выполнении условий настоящего договора, доверять, согласовывать все свои действия, принимать и следовать предлагаемым рекомендациям и разъяснениям.";
    var p29 = docx.createP();
    p29.addText(t29,{font_face: 'Times New Roman', font_size: '12'});

    const t30 = "2.2.2. Оплатить услуги Исполнителя в размере и сроки, указанные в настоящем договоре.";
    var p30 = docx.createP();
    p30.addText(t30,{font_face: 'Times New Roman', font_size: '12'});

    const t31 = "2.2.3. Предоставить Исполнителю по акту приема - передачи все необходимые, имеющиеся документы по делу о признании договора займа с микрофинансовой организацией недействительным в оригиналах и копиях, для установления предмета и основания иска, подготовки искового заявления и последующей передачи в судебные органы.";
    var p31 = docx.createP();
    p31.addText(t31,{font_face: 'Times New Roman', font_size: '12'});

    const t32 = "2.2.4. Предоставить Исполнителю достоверную информацию и оказывать содействие в получении информации, которой владеет Заказчик имеющей существенное значение для правильного и всестороннего рассмотрения дела.";
    var p32 = docx.createP();
    p32.addText(t32,{font_face: 'Times New Roman', font_size: '12'});

    const t33 = "2.2.5. Предоставить Исполнителю доверенность, с правом передоверия, на ведение судебного дела и с правом получения, всех необходимых документов, которые могут понадобиться в ходе исполнения условий настоящего договора.";
    var p33 = docx.createP();
    p33.addText(t33,{font_face: 'Times New Roman', font_size: '12'});

    const title5 = "3. Расчеты и порядок оплаты";
    var title5P = docx.createP({align: 'center'});
    title5P.addText(title5,{font_face: 'Times New Roman', bold: true, font_size: '12'});

    const t34 = "3.1. За выполнение юридических услуг, указанных в настоящем договоре, Заказчик выплачивает Исполнителю денежную сумму вознаграждения  путем перечисления не позднее 3-х дней  с момента подписания настоящего договора.";
    var p34 = docx.createP();
    p34.addText(t34,{font_face: 'Times New Roman', font_size: '12'});

    const t35 = "3.2. Стоимость услуг и порядок оплаты регулируется Приложением №1 к настоящему договору .";
    var p35 = docx.createP();
    p35.addText(t35,{font_face: 'Times New Roman', font_size: '12'});

    const t36 = "3.3. Произведенные расчеты по оплате по данному Договору, являются независимыми от результата и исхода дела. При этом, произведенные расчеты, при окончании или необоснованном расторжении Договора, считаются полностью отработанными.";
    var p36 = docx.createP();
    p36.addText(t36,{font_face: 'Times New Roman', font_size: '12'});

    const title6 = "4. Ответственность сторон и порядок расторжения договора";
    var title6P = docx.createP({align: 'center'});
    title6P.addText(title6,{font_face: 'Times New Roman', bold: true, font_size: '12'});

    const t37 = "4.1. Стороны примут меры к разрешению всех споров и разногласий, которые могут возникнуть  из настоящего договора или из отдельных его пунктов, путем переговоров.";
    var p37 = docx.createP();
    p37.addText(t37,{font_face: 'Times New Roman', font_size: '12'});

    const t38 = "4.2. В случае, если стороны не придут к соглашению, все споры и разногласия разрешаются в соответствии с действующим законодательством Республики Казахстан.";
    var p38 = docx.createP();
    p38.addText(t38,{font_face: 'Times New Roman', font_size: '12'});

    const t39 = "4.3. Стороны соглашаются с договорной подсудностью. Все возникшие споры между сторонами по настоящему договору, будут рассматриваться в Специализированном межрайонном экономическом суде г.Астаны.";
    var p39 = docx.createP();
    p39.addText(t39,{font_face: 'Times New Roman', font_size: '12'});

    const t40 = "4.4. Меры ответственности сторон, не предусмотренные в настоящем Договоре, применяются в соответствии с нормами гражданского законодательства, действующего на территории Республики Казахстан.";
    var p40 = docx.createP();
    p40.addText(t40,{font_face: 'Times New Roman', font_size: '12'});

    const t41 = "4.5. В случае неисполнения Заказчиком п.2.2.5. настоящего Договора, ответственность за наступление отрицательных последствий несет Заказчик.";
    var p41 = docx.createP();
    p41.addText(t41,{font_face: 'Times New Roman', font_size: '12'});

    const t42 = "4.6. В случае неисполнения исполнителем п.1.2 и 1.3 настоящего Договора, повлекших отрицательные последствия, ответственность несет исполнитель при наличии вины.";
    var p42 = docx.createP();
    p42.addText(t42,{font_face: 'Times New Roman', font_size: '12'});

    const title7 = "5. Действие договора, заключительные положения  и прочие условия";
    var title7P = docx.createP({align: 'center'});
    title7P.addText(title7,{font_face: 'Times New Roman', bold: true, font_size: '12'});
    
    const t43 = "5.1. Настоящий договор вступает в силу с момента его подписания сторонами и действует до полного исполнения обязательства, по оказанию юридических услуг.";
    var p43 = docx.createP();
    p43.addText(t43,{font_face: 'Times New Roman', font_size: '12'});

    const t44 = "5.2. Любые изменения и дополнения к настоящему Договору действительны лишь при условии, что они совершены в письменной форме и подписаны сторонами.";
    var p44 = docx.createP();
    p44.addText(t44,{font_face: 'Times New Roman', font_size: '12'});

    const t45 = "5.3. В случае расторжения настоящего договора по инициативе Заказчика, денежные средства, указанные в п.3.1. настоящего договора - не возвращаются.";
    var p45 = docx.createP();
    p45.addText(t45,{font_face: 'Times New Roman', font_size: '12'});

    const t46 = "5.4. Настоящий Договор составлен в соответствии с гражданским законодательством Республики Казахстан, в 2-х экземплярах на русском языке. Оба экземпляра идентичны и имеют одинаковую силу. У каждой из сторон находится один экземпляр настоящего Договора.";
    var p46 = docx.createP();
    p46.addText(t46,{font_face: 'Times New Roman', font_size: '12'});

    const t47 = "Юридический адрес сторон и банковские реквизиты";
    var p47 = docx.createP({align: 'center'});
    p47.addText(t47,{font_face: 'Times New Roman', font_size: '12'});
    splittedName = var3.split(" ");
    if(splittedName[2] === undefined){
        splittedName[2] = "  ";
    }
    /* table 1 */
    var table1 = [
        [{
          val: "Исполнитель",
          opts: {
            cellColWidth: 4261,
            b:true,
            shd: {
              themeFill: "text1",
              "themeFillTint": "80"
            },
          }
        },{
          val: "Заказчик",
          opts: {
            cellColWidth: 4261,
            b:true,
            shd: {
              themeFill: "text1",
              "themeFillTint": "80"
            }
          }
        }],
        ['ТОО «GEN4»\nБИН: 150240008616\nБанк получатель: AO Kaspi Bank\nБИК: CASPKZKA\nНомер счета:\nKZ46722S000001768682\nКБе 17\nІскендір Нұрхан Қонысбайұлы\nТел: +7-747-676-63-63\n\n\n\n\n\n\n\n\n\n\n\nДиректор Іскедір Н.Қ.____________',"ФИО : " + var3+"\n"+"ИИН: "+var4 + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\nФИО: "+splittedName[0]+" "+splittedName[1][0]+"."+splittedName[2][0]+'.____________'],
      ]
       
      var tableStyle = {
        tableSize: 24,
        tableColor: "ada",
        tableAlign: "left",
        tableFontFamily: "Times New Roman",
        borders: true,
        newline: '\n'
      }
       
      docx.createTable (table1, tableStyle);

    /* table 1 ends */

    
    const title8 = "  ПРИЛОЖЕНИЕ № 1";
    var title8P = docx.createP({align: 'center'});
    title8P.addText(title8, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    const title9 = "К ДОГОВОРУ № "+var1;
    var title9P = docx.createP({align: 'center'});
    title9P.addText(title9, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    const t48 = "ТОО «GEN4» в лице Директора – Іскендір Нұрхан Қонысбайұлы, действующего на основании Устава, именуемое в дальнейшем «Исполнитель» с одной стороны и "+var3+" Заказчика именуемый в дальнейшем «Заказчик» с другой стороны, заключили настоящее Приложение к Договору об оказании юридических услуг (далее – Приложение), о нижеследующем:";
    var p48 = docx.createP();
    p48.addText(t48,{font_face: 'Times New Roman', font_size: '12'});

    const t49 = "1.ПЕРЕЧЕНЬ УСЛУГ ОКАЗЫВАЕМЫХ ЗАКАЗЧИКУ";
    var p49 = docx.createP();
    p49.addText(t49, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    organizations = "";
    col = 1;
    if(Array.isArray(var6)){
        for(let i = 0; i < var6.length; i++){
            organizations += var6[i]+",";
        }
        col = var6.length;
    }else{
        organizations = var6;
    }
    organizations = organizations.substring(0, organizations.length - 1);

    const t50 = "1.Перечень услуг оказываемых заказчику в рамках оказания услуги, исполнитель берет на себя обязательства по признанию договоров недействительными между заказчиком и "+organizations+" ("+col+") , (далее – Кредитор). В рамках оказания услуги, исполнитель берет на себя обязательства:";
    var p50 = docx.createP();
    p50.addText(t50,{font_face: 'Times New Roman', font_size: '12'});

    const t51 = "А) по досудебному урегулированию; ";
    var p51 = docx.createP();
    p51.addText(t51,{font_face: 'Times New Roman', font_size: '12'});

    const t52 = "Б) представительству в судах первой инстанции;";
    var p52 = docx.createP();
    p52.addText(t52,{font_face: 'Times New Roman', font_size: '12'});

    const t53 = "В) отмене ранее вынесенному судебному решению;";
    var p53 = docx.createP();
    p53.addText(t53,{font_face: 'Times New Roman', font_size: '12'});

    const t54 = "Г) отмене ранее вынесенной исполнительской надписи.";
    var p54 = docx.createP();
    p54.addText(t54,{font_face: 'Times New Roman', font_size: '12'});

    const t55 = "2.  СПИСОК ОРГАНИЗАЦИЙ И ДОГОВОРОВ, А ТАКЖЕ СТОИМОСТЬ УСЛУГ, ПО КОТОРЫМ ОКАЗЫВАЕТСЯ ЮРИДИЧЕСКАЯ ПОМОЩЬ ИСПОЛНИТЕЛЕМ ЗАКАЗЧИКУ";
    var p55 = docx.createP();
    p55.addText(t55, {bold: true,font_face: 'Times New Roman', font_size: '12'});
    /* table 2 */
    var table2 = [
        [{
          val: "№ п/п",
          opts: {
            cellColWidth: 1700,
            b:true,
            align: "center",
            shd: {
              themeFill: "text1",
              "themeFillTint": "80"
            },
          }
        },{
            val: "Наименование организации – кредитора",
            opts: {
              cellColWidth: 1700,
              b:true,
              align: "center",
              shd: {
                themeFill: "text1",
                "themeFillTint": "80"
              }
            }
          },{
            val: "Номер договора займа",
            opts: {
              cellColWidth: 1700,
              b:true,
              align: "center",
              shd: {
                themeFill: "text1",
                "themeFillTint": "80"
              }
            }
          },{
            val: "Дата договора займа",
            opts: {
              cellColWidth: 1700,
              b:true,
              align: "center",
              shd: {
                themeFill: "text1",
                "themeFillTint": "80"
              }
            }
          },{
            val: "Стоимость услуг Заказчика",
            opts: {
              cellColWidth: 1700,
              b:true,
              align: "center",
              shd: {
                themeFill: "text1",
                "themeFillTint": "80"
              }
            }
          }],
      ]
      if(Array.isArray(var6)){
        for(let i = 0; i < var6.length; i++){    
            table2.push([i + 1, var6[i], var7[i], var8[i], var9[i]]);
          }
      }else{
          table2.push([1, var6, var7, var8, var9]);
      }
      var tableStyle = {
        tableSize: 12,
        tableAlign: "center",
        tableFontFamily: "Times New Roman",
        borders: true,
      }
       
      docx.createTable (table2, tableStyle);

    /* table 2 ends */

    const t56 = "3. СТОИМОСТЬ УСЛУГ И УСЛОВИЯ ОПЛАТЫ";
    var p56 = docx.createP();
    p56.addText(t56, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    const t57 = "3.1. Стоимость услуг составляет "+var10+' тенге ('+var11+' тенге)';
    var p57 = docx.createP();
    p57.addText(t57,{font_face: 'Times New Roman', font_size: '12'});

    const t58 = "3.2. Порядок оплаты: предоплата 50% и окончательный платеж 50%.";
    var p58 = docx.createP();
    p58.addText(t58,{font_face: 'Times New Roman', font_size: '12'});

    const t59 = "3.3. Заказчик производит 50% - ю предоплату стоимости услуг в течение 3 (треx) рабочих дней после подписания настоящего договора.";
    var p59 = docx.createP();
    p59.addText(t59,{font_face: 'Times New Roman', font_size: '12'});

    const t60 = "3.4. Заказчик производит  окончательный платеж в размере 50% от стоимости услуг в течение  	3 (треx) рабочих дней с момента получения решения суда.";
    var p60 = docx.createP();
    p60.addText(t60,{font_face: 'Times New Roman', font_size: '12'});

    const t61 = "3.5.В случае не исполнения заказчиком своих обязательств, указанных в п.2.4. Исполнитель вправе начислить единовременный штраф в размере 10 мрп.";
    var p61 = docx.createP();
    p61.addText(t61,{font_face: 'Times New Roman', font_size: '12'});

    const t62 = "3.6. Расходы, фактически понесенные исполнителем при исполнении данного договора, подлежат оплате заказчиком в десятидневный срок с момента предоставления подтверждающих документов исполнителем, с дальнейшим возмещением с контрагента  заказчика в судебном порядке.";
    var p62 = docx.createP();
    p62.addText(t62,{font_face: 'Times New Roman', font_size: '12'});

    const t63 = " 4. ПРОЧИЕ УСЛОВИЯ";
    var p63 = docx.createP();
    p63.addText(t63,{font_face: 'Times New Roman', font_size: '12'});

    const t64 = "4.1. Во всем остальном, что не предусмотрено настоящим Приложением, действуют условия Договора.";
    var p64 = docx.createP();
    p64.addText(t64,{font_face: 'Times New Roman', font_size: '12'});

    const t65 = "4.2. Договор действует с момента подписания, до 31.06.2019 года, а в части взаиморасчетов до полного их исполнения.";
    var p65 = docx.createP();
    p65.addText(t65,{font_face: 'Times New Roman', font_size: '12'});

    const t66 = "4.3 Под надлежащим исполнением стороны понимают вступившее в законную силу решение суда, согласно которому договоры между Заказчиком и его Кредиторами будут признаны недействительными.";
    var p66 = docx.createP();
    p66.addText(t66,{font_face: 'Times New Roman', font_size: '12'});

    const t67 = "5. ПОДПИСИ СТОРОН:";
    var p67 = docx.createP();
    p67.addText(t67, {bold: true,font_face: 'Times New Roman', font_size: '12'});

    const t68 = "Юридический адрес сторон и банковские реквизиты";
    var p68 = docx.createP({align: 'center'});
    p68.addText(t68, {bold: true, underline: true,font_face: 'Times New Roman', font_size: '12'});

      /* table 1 */
    var table3 = [
        [{
          val: "Исполнитель",
          opts: {
            cellColWidth: 4261,
            b:true,
            shd: {
              themeFill: "text1",
              "themeFillTint": "80"
            },
          }
        },{
          val: "Заказчик",
          opts: {
            cellColWidth: 4261,
            b:true,
            shd: {
              themeFill: "text1",
              "themeFillTint": "80"
            }
          }
        }],
        ['ТОО «GEN4»\nБИН: 150240008616\nБанк получатель: AO Kaspi Bank\nБИК: CASPKZKA\nНомер счета:\nKZ46722S000001768682\nКБе 17\nІскендір Нұрхан Қонысбайұлы\nТел: +7-747-676-63-63\n\n\n\n\n\n\n\n\n\n\n\nДиректор Іскедір Н.Қ.____________',"ФИО : " + var3+"\n"+"ИИН: "+var4 + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\nФИО: "+splittedName[0]+" "+splittedName[1][0]+"."+splittedName[2][0]+'.____________'],
      ]
       
      var tableStyle = {
        tableSize: 24,
        tableColor: "ada",
        tableAlign: "left",
        tableFontFamily: "Times New Roman",
        borders: true
      }
       
      docx.createTable (table3, tableStyle);

    /* table 3 ends */

    docx.generate(res);
});
////
/////
app.get('/claimtemplate', function(req, res){
    let docx = officegen('docx');
    const { Document, Paragraph, Packer } = docx;
    var tempFilePath = tempfile('.docx');
    docx.setDocSubject ( 'testDoc Subject' );
    docx.setDocKeywords ( 'keywords' );
    docx.setDescription ( 'test description' );
    
    docx.on('finalize', function(written) {
        console.log('Finish to create Word file.\nTotal bytes created: ' + written + '\n');
        
    });
    docx.on('error', function(err) {
        console.log(err);
    });
    var datetime = new Date();
   res.writeHead ( 200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.documentml.document",
    'Content-disposition': 'attachment; filename=Dosudebnaya%20Pretenziya.docx'
    });

    var var1 = req.query[1], var2 = req.query[2], var3 = req.query[3], var4 = req.query[4], var5 = req.query[5], var6 = req.query[6], var7 = req.query[7], var8 = req.query[8], var9 = req.query[9], var10 = req.query[10], var11 = req.query[11], var12 = req.query[12], var13 = req.query[13], var14 = req.query[14], var15 = req.query[15], var16 = req.query[16], var17=req.query[17], var18=req.query[18], var19=req.query[19], var20=req.query[20];
    var datetime = new Date();
    let datestring = ""+datetime.getDate()+"."+(datetime.getMonth() + 1)+"."+datetime.getFullYear().toString();
    
    const t1 = var1;
    var p1 = docx.createP({align: 'right'});
    p1.addText(t1, {font_face: "Times New Roman", font_size: '14'});

    const t2 = var2;
    var p2 = docx.createP({align: 'right'});
    p2.addText(t2, {font_face: "Times New Roman", font_size: '14'});

    const t3 = var3;
    var p3 = docx.createP({align: 'right'});
    p3.addText(t3, {font_face: "Times New Roman", font_size: '14'});

    const t4 = "Реквизиты:";
    var p4 = docx.createP({align: 'right'});
    p4.addText(t4, {font_face: "Times New Roman", font_size: '14'});

    const t5 = "БИН:"+var4;
    var p5 = docx.createP({align: 'right'});
    p5.addText(t5, {font_face: "Times New Roman", font_size: '14'});

    const t6 = "ИИК:"+var5;
    var p6 = docx.createP({align: 'right'});
    p6.addText(t6, {font_face: "Times New Roman", font_size: '14'});

    const t7 = "БИК"+var6;
    var p7 = docx.createP({align: 'right'});
    p7.addText(t7, {font_face: "Times New Roman", font_size: '14'});

    const t8 = var7;
    var p8 = docx.createP({align: 'right'});
    p8.addText(t8, {font_face: "Times New Roman", font_size: '14'});

    const t9 = "email: "+var8;
    var p9 = docx.createP({align: 'right'});
    p9.addText(t9, {font_face: "Times New Roman", font_size: '14'});

    const t10 = "тел: "+var9;
    var p10 = docx.createP({align: 'right'});
    p10.addText(t10, {font_face: "Times New Roman", font_size: '14'});

    var emptyP = docx.createP();

    const t11 = "Должник: "+var10;
    var p11 = docx.createP({align: 'right'});
    p11.addText(t11, {font_face: "Times New Roman", font_size: '14'});

    const t12 = var11;
    var p12 = docx.createP({align: 'right'});
    p12.addText(t12, {font_face: "Times New Roman", font_size: '14'});

    const t13 = "ИИН: "+var12;
    var p13 = docx.createP({align: 'right'});
    p13.addText(t13, {font_face: "Times New Roman", font_size: '14'});

    const t14 = "Тел: "+var13;
    var p14 = docx.createP({align: 'right'});
    p14.addText(t14, {font_face: "Times New Roman", font_size: '14'});

    const t15 = "Исковое заявление";
    var p15 = docx.createP({align: 'center'});
    p15.addText(t15, {font_face: "Times New Roman", font_size: '14'});

    const t16 = "Должник проживает в принадлежащей ей на праве собственности квартире № "+var14+", расположенной в доме 230, по улице Жарокова, в городе Алматы, который обслуживает "+var2+", согласно Устава и решения собрания участников.";
    var p16 = docx.createP();
    p16.addText(t16, {font_face: "Times New Roman", font_size: '14'});

    const t17 = "Согласно требованиям ч.3 ст.18 Закона РК «О жилищных отношениях» собственники помещении (квартир), входящих в состав объекта кондоминиума, также несут обязанности, предусмотренные статьями 35 и 50 настоящего закона.";
    var p17 = docx.createP();
    p17.addText(t17, {font_face: "Times New Roman", font_size: '14'});

    const t18 = "Согласно ч.8 ст. 43 Закона РК «О жилищных отношениях» собственники помещений не участвующие в управлении делами кооператива, наряду со всеми членами кооператива обязаны принимать соразмерное денежное и (или) трудовое участие в содержании объекта кондоминиума.";
    var p18 = docx.createP();
    p18.addText(t18, {font_face: "Times New Roman", font_size: '14'});

    const t19 = "В соответствии со ст. 50 Закона РК «О жилищных отношениях» собственники помещений (квартир) обязаны участвовать в расходах на содержание общего имущества объекта кондоминиума. Расходы на содержание общего имущества объекта кондоминиума производятся ежемесячно. Размеры расходов на содержание общего имущества объекта кондоминиума устанавливаются соразмерно доле собственника помещения (квартиры) в общем имуществе.";
    var p19 = docx.createP();
    p19.addText(t19, {font_face: "Times New Roman", font_size: '14'});

    const t20 = "Однако должник в течение длительного времени не исполняет свои обязанности по оплате взносов на содержание жилья. Несмотря на предупреждения администрации "+var2+", должник не оплачивает взносы и задолженность за ним составила "+var15+" ("+var18+") тенге. Согласно протокола общего собрания от 17 октября 2016 года тариф составляет 100 тенге за 1 кв.м. Согласно протокола общего собрания от 17 октября 2016 года тариф на накопление на капитальный ремонт составляет 10 тенге за 1 кв.м.  Согласно протокола общего собрания от 17 октября 2016 года тариф на содержание за 1 единицу автомобильного  места составляет 4000 тенге . Расчет задолженности прилагается.";
    var p20 = docx.createP();
    p20.addText(t20, {font_face: "Times New Roman", font_size: '14'});

    const t21 = "В соответствий с п. 4 ст. 50 Закона РК «О жилищных отношениях» На требование по погашению задолженности срок исковой давности не распространяется.";
    var p21 = docx.createP()
    p21.addText(t21, {font_face: "Times New Roman", font_size: '14'});

    const t22 = "Согласно п.п. 10, п. 1 ст. 535 Налогового Кодекса РК с заявлении о вынесении судебного приказа – 50 процентов от ставок государственной пошлины.";
    var p22 = docx.createP();
    p22.addText(t22, {font_face: "Times New Roman", font_size: '14'});

    const t23 = "Таким образом, в связи с невозможностью мирного разрешения спора, после неоднократных уведомлении и нежеланием должника добровольно погасить задолженность, образовавшуюся за период проживания и пользования жилищем. ";
    var p23 = docx.createP();
    p23.addText(t23, {font_face: "Times New Roman", font_size: '14'});

    const t24 = "На основании изложенного, руководствуясь ст.ст. 18, 43, 50 Закона РК «О жилищных отношениях», ст.ст. 148, 149 ГПК РК;";
    var p24 = docx.createP();
    p24.addText(t24, {font_face: "Times New Roman", font_size: '14'});

    const t25 = "Прошу суд:";
    var p25 = docx.createP({align: 'center'});
    p25.addText(t25, {font_face: "Times New Roman", font_size: '14'});

    const t26 = "Взыскать с должника Диас Дины сумму основного долга в размере "+var15+" ("+var18+") тенге, сумму государственной пошлины в размере "+var16+" ("+var19+") тенге.";
    var p26 = docx.createP();
    p26.addText(t26, {font_face: "Times New Roman", font_size: '14'});
    
    const t27 = "Всего: "+var17+" ("+var20+") тенге.";
    var p27 = docx.createP();
    p27.addText(t27, {font_face: "Times New Roman", font_size: '14'});

    const t28 = "Приложение:";
    var p28 = docx.createP();
    p28.addText(t28, {font_face: "Times New Roman", font_size: '14'});

    const t29 = "1. Копия искового заявления";
    var p29 = docx.createP();
    p29.addText(t29, {font_face: "Times New Roman", font_size: '14'});

    const t30 = "2. Копия устава истца";
    var p30 = docx.createP();
    p30.addText(t30, {font_face: "Times New Roman", font_size: '14'});

    const t31 = "3. Квитанция об оплате государственной пошлины";
    var p31 = docx.createP();
    p31.addText(t31, {font_face: "Times New Roman", font_size: '14'});

    const t32 = "4. Копия протокола общего собрания";
    var p32 = docx.createP();
    p32.addText(t32, {font_face: "Times New Roman", font_size: '14'});

    const t33 = "5. Расчет образовавшейся задолженности";
    var p33 = docx.createP();
    p33.addText(t33, {font_face: "Times New Roman", font_size: '14'});

    const t34 = "6. Копия доверенности";
    var p34 = docx.createP();
    p34.addText(t34, {font_face: "Times New Roman", font_size: '14'});

    const t35 = "7. Решение единственного участника";
    var p35 = docx.createP();
    p35.addText(t35, {font_face: "Times New Roman", font_size: '14'});

    const t36 = "8. Приказ о вступлении на должность директора";
    var p36 = docx.createP();
    p36.addText(t36, {font_face: "Times New Roman", font_size: '14'});

    const t37 = "9. Копия удостоверения личности на Токбаева Д.Е.";
    var p37 = docx.createP();
    p37.addText(t37, {font_face: "Times New Roman", font_size: '14'});

    const t38 = "10. Копия удостоверения личности на Калдыбаева Ч.К.";
    var p38 = docx.createP();
    p38.addText(t38, {font_face: "Times New Roman", font_size: '14'});

    const t39 = "11. Копия диплома";
    var p39 = docx.createP();
    p39.addText(t39, {font_face: "Times New Roman", font_size: '14'});

    const t40 = "12. Копия досудебной претензии";
    var p40 = docx.createP();
    p40.addText(t40, {font_face: "Times New Roman", font_size: '14'});

    const t41 = "13. Уведомление о направлении досудебной претензии";
    var p41 = docx.createP();
    p41.addText(t41, {font_face: "Times New Roman", font_size: '14'});

    const t42 = "14. Договор на оказание услуг по взысканию дебиторской задолженности";
    var p42 = docx.createP();
    p42.addText(t42, {font_face: "Times New Roman", font_size: '14'});

    const t43 = "15. Приказ о приеме на работу";
    var p43 = docx.createP();
    p43.addText(t43, {font_face: "Times New Roman", font_size: '14'});

    const t44 = "Представитель";
    var p44 = docx.createP();
    p44.addText(t44, {font_face: "Times New Roman", font_size: '14'});

    const t45 = var2+"                            Калдыбаев Ч.К.";
    var p45 = docx.createP();
    p45.addText(t45, {font_face: "Times New Roman", font_size: '14'});


    docx.generate(res);
});
/////
app.listen(8080, ()=>{
    console.log("listening to the port 8080");
})