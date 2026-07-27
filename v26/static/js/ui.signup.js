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
$("#signup").on('click',function(){
    obi = [g("user").value,g("password").value,g("email").value]
    console.log(obi)
    if (!(checkUnd(obi))){
        data = {
            "user":g("user").value,
            "password":g("password").value,
            "email":g("email").value

    }
    data = JSON.stringify(data);
    const fd = new FormData();
    // append directly as part of the postData in plain text
    fd.append('data', data);
    console.log(data)
    fetch("/signupac",
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
            Swal.fire({
                title: "Hey!",
                text:  data["output"].toString() ,
                icon: "info"
              });
        })
    }else{
        Swal.fire({
            title: "Error!",
            text: g("error01").innerHTML.toString(),
            icon: "error"
          });
    }
})