// UISIGN
function g(id){
    return document.getElementById(id)
}
function checkUnd(list){
    if (list.includes(undefined) || list.includes("")){
        return true
    }else{
        return false
    }
}


function setCookie(name,value,days) {
    var expires = "";
    if (days) {
        var date = new Date();
        date.setTime(date.getTime() + (days*24*60*60*1000));
        expires = "; expires=" + date.toUTCString();
    }
    document.cookie = name + "=" + (value || "")  + expires + "; path=/";
}
function getCookie(name) {
    var nameEQ = name + "=";
    var ca = document.cookie.split(';');
    for(var i=0;i < ca.length;i++) {
        var c = ca[i];
        while (c.charAt(0)==' ') c = c.substring(1,c.length);
        if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
    }
    return null;
}
function eraseCookie(name) {   
    document.cookie = name +'=; Path=/; Expires=Thu, 01 Jan 1970 00:00:01 GMT;';
}
$("#signin").on('click',function(){
    fetch('https://api.ipify.org?format=json')
    .then(response => response.json())
    .then(data => {
        obi = [g("user").value,g("password").value]
        console.log(obi)
        if (!(checkUnd(obi))){
            data = {
                "user":g("user").value,
                "ip":data.ip,
                "password":g("password").value

        }
        data = JSON.stringify(data);
        const fd = new FormData();
        fd.append('data', data);
        console.log(data)
        fetch("/signinac",
            {
                method: "POST",
                // headers: {
                //   'Content-Type': 'charset=utf-8'
                //   'Content-Type': 'multipart/form-data; charset=utf-8'
                // },
                body: fd
            })
            .then(function(res){ return res.json(); })
            .then(function(data){ 
                console.log(data)
                if (data["output"].toString()=="success"){
                    console.log("Hi")
                    setCookie("_ie",data["cookie"],7)
                    Swal.fire({
                        title: "Hey!",
                        text:  data["message"].toString() ,
                        icon: "success"
                    });
                    setTimeout(() => {
                        
                        window.location.href = window.location.origin + "/"
                    }, (1500));
                }else{

                    Swal.fire({
                        title: "Hey!",
                        text:  data["output"].toString() ,
                        icon: "info"
                    });

                }
            })
        }else{
            Swal.fire({
                title: "Error!",
                text: g("error01").innerHTML.toString(),
                icon: "error"
            });
        }    
    })
    .catch(error => {
        console.log('Error:', error);
    });
})