// JavaScript Document
	function showTime(){
		myDate = new Date();
		year = myDate.getYear();
		day=myDate.getDate();
		hours=myDate.getHours();
	    minutes=myDate.getMinutes();
		seconds=myDate.getSeconds();
		if (year < 1000)
		year+=1900
		if (day<10)
		day="0"+day
		if (hours<=9)
		hours="0"+hours
		if (minutes<=9)
		minutes="0"+minutes
		if (seconds<=9)
		seconds="0"+seconds
		var dTimeMHMD ="<font size='2' color='#406480'>" + hours + ":" + minutes + ":" + seconds + "</font>";
		showMonth();
		var dDateMHMD ="<font size='2' color='#406480'>" + year + "-" + myDate.month +  "-" + day + "</font>";
		
		
		
		if (document.all)
			document.all.dDate.innerHTML=dDateMHMD
		else if (document.getElementById)
			document.getElementById("dDate").innerHTML=dDateMHMD
		else
		document.write(dDate)

		if (document.all)
			document.all.dTime.innerHTML=dTimeMHMD
		else if (document.getElementById)
			document.getElementById("dTime").innerHTML=dTimeMHMD
		else
		document.write(dTime)
	}
	
	

	
	setInterval("showTime()",1000);
	
	function showMonth(){	
		switch(myDate.getMonth()){
			case 0 :  myDate.month = "Jan";
				break;
			case 1 : myDate.month = "Feb";
				break;
			case 2 : myDate.month = "Mar";
				break;
			case 3 : myDate.month = "Apr";
				break;
			case 4 : myDate.month = "May";
				break;
			case 5 : myDate.month = "Jun";
				break;
			case 6 : myDate.month = "Jul";
				break;
			case 7 : myDate.month = "Aug";
				break;
			case 8 : myDate.month = "Sep";
				break;
			case 9 : myDate.month = "Oct";
				break;
			case 10 : myDate.month = "Nov";
				break;
			case 11 : myDate.month = "Dec";
				break;
			default: null;	
			}
		}
		
	function search1(){
		window.open("http://www.google.com/search?sourceid=navclient&ie=UTF-8&q="+ document.frm.gSearch.value, "_blank", "");
	}
	
	function login(){
		if(document.frm1.user.value == "admin" && document.frm1.pass.value == "javascript")
			window.open("user_settings.html", "_self");
		else
			window.alert("Username or Password is wrong. plz try again!!");
	}
	
	function checkInfo(){
		var status = true;
		if(document.frm2.fName.value.length == 0){
			first.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'First Name' field."+"</font>";
			status = false;
		}
		else {
			first.innerHTML = "";
			status = true;
		}
		
		if(document.frm2.lName.value.length == 0){
			last.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'Last Name' field."+"</font>";
			statue = false;
		}
		else {
			last.innerHTML = "";
			status = true;
		}
		
		if(document.frm2.uName.value.length == 0){
			user.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'User Name' field."+"</font>";
			status = false;
		}
		else {
			user.innerHTML = "";
			status = true;
		}
		
		if(document.frm2.pass.value.length == 0){
			pas.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'Password' field."+"</font>";
			status = false;
		}
		
		else if(document.frm2.pass.value.length != 0 && document.frm2.pass.value.length <= 6 ){
			pas.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Password must be more than 6 character."+"</font>";
			status = false;
		}
		else {
			pas.innerHTML = "";
			status = true;
		}
		
		if(document.frm2.cPass.value.length == 0){
			cpas.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'Confirm Password' field."+"</font>";
			status = false;
		}
		
		else if(document.frm2.cPass.value.length != 0 && document.frm2.cPass.value != document.frm2.pass.value){
			cpas.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Password and Confirm Password aren't match."+"</font>";
			status = false;
		}
		else {
			cpas.innerHTML = "";
			status = true;
		}
		
		if(document.frm2.email.value.length == 0){
			mail.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'E-mail' field."+"</font>";
			status = false;
		}
		else if(document.frm2.email.value.length != 0 && document.frm2.email.value.indexOf("@") == -1){
			mail.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please fill 'E-mail' correctly!."+"</font>";
			status = false;
		}
		else {
			mail.innerHTML = "";
			status = true;
		}
		
		if((document.frm2.gender[0].checked=="") && (document.frm2.gender[1].checked=="")){
			gen.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please select your sex"+"</font>";
			status = false;
		}
		else {
			gen.innerHTML = "";
			status = true;
		}
				
		if((document.frm2.Sport.checked=="")&&(document.frm2.Movie.checked=="")&&(document.frm2.Music.checked=="")){
			hob.innerHTML = "<font face='verdana' size='1' color='pink'>"+"*Please select your hobby."+"</font>";
			status = false;
		}
		else {
			hob.innerHTML = "";
			status = true;
		}
		
		if(status == false)
			return false;
		else
			return true;
		
		
	}
	
	function ld(){
		document.frm2.fName.value="";
		document.frm2.lName.value="";
		document.frm2.uName.value="";
		document.frm2.pass.value="";
 		document.frm2.email.value="";
		document.frm2.cPass.value="";
		document.frm2.gender[0].checked=false;
		document.frm2.gender[1].checked=false;
		document.frm2.Movie.checked==false;
		document.frm2.Sport.checked==false;
		document.frm2.Music.checked==false;
	}
	
	function subs(head){
		if(head.style.display == "none")
			head.style.display = "";
		else
			head.style.display = "none";
	}
	
	function func(){
		if(document.set.selectu.value == "color"){
			document.set.chColor.disabled = false;
			document.set.chBG.disabled = true;
			}
		else if(document.set.selectu.value == "bg"){
			document.set.chBG.disabled = false;
			document.set.chColor.disabled = true;
			}
		else {
			document.set.chColor.disabled = true;
			document.set.chBG.disabled = true;
		}
	}
	
	function colorFunc(){
		var i = document.set.chColor.selectedIndex;
		bgColor = document.set.chColor.options[i].value;
		document.body.style.backgroundImage ="";	
		document.body.style.backgroundColor = bgColor;
	}
	
	function bgFunc(){
		var i = document.set.chBG.selectedIndex;
		bgImage = document.set.chBG.options[i].value;	
		document.body.style.backgroundImage = bgImage;
	}
	
	function changeScrollbarColor(C){
    	document.body.style.scrollbarBaseColor = C ; 
	}
	
	function defaultScroll(){
		document.body.style.scrollbarBaseColor = "";
	}
	
var second = null;
var tick = null;
 
function timer( ){
    second = -1;
    tick = setInterval("ticktack( )", 1000);
    ticktack( );
	}
 

	function ticktack( ){
    ++second;
    var sec = second;
    var hr = Math.floor( sec / 3600 );
    sec %= 3600;
    var mn = Math.floor( sec / 60 );
    sec %= 60;
    var show = ( hr < 10 ? "0" : "" ) + hr
               + ":" + ( mn< 10 ? "0" : "" ) + mn
               + ":" + ( sec < 10 ? "0" : "" ) + sec;
    document.getElementById("elapsed").innerHTML = show;
	}
	
