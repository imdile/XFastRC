async function addPeople() {
  var data = await Swal.fire({
    title: "Add user",
    html: `
<style>
.addpeople-form *{
    float:unset !important;
}
.form-check .form-check-input{
    float: left !important;
}

.form-check .form-check-label{
    float: left !important;
    margin-bottom: 4px;
    margin-top : 4px
}

.swal2-input{
  margin:10px;
  margin-left: 20px;
  width:100%;
}
</style>
<div class="addpeople-form" style="height: 100%;">
  <div class="maincon" style="float: left !important;width:43%; ">
    <h4>Username:</h4>
    <input id="swal" class="swal2-input" placeholder="Username">
  </div>
  <div class="maincon" style="float: right !important;width:40%;">
    <h4 style="margin-bottom: 17px">Permission:</h4>
    <div class="form-check">
      <input class="form-check-input" type="checkbox" value="" id="checkbox1">
      <label class="form-check-label" for="checkbox1">
        checkbox1
      </label>
    </div>
    <div class="form-check">
      <input class="form-check-input" type="checkbox" value="" id="checkbox2">
      <label class="form-check-label" for="checkbox2">
        checkbox2
      </label>
    </div>
    <div class="form-check">
      <input class="form-check-input" type="checkbox" value="" id="checkbox3">
      <label class="form-check-label" for="checkbox3">
        checkbox3
      </label>
    </div>

  </div>
  <br>
  <div>
    <h4 style="margin-bottom: 17px">Permission:</h4>

  <select id="swal2-select" style="margin-bottom: 3px;margin-top: 15px;margin-left: 15%;margin-right: 17%;" class="swal2-select"><option value="" disabled="">Select a fruit</option><optgroup label="Fruits"><option value="apples">Apples</option><option value="bananas">Bananas</option><option value="grapes">Grapes</option><option value="oranges">Oranges</option></optgroup><optgroup label="Vegetables"><option value="potato">Potato</option><option value="broccoli">Broccoli</option><option value="carrot">Carrot</option></optgroup><option value="icecream">Ice cream</option></select>
</div>
  </div>
      `,
    focusConfirm: false,
    preConfirm: () => {
      return [
        document.getElementById("swal-input1").value,
        document.getElementById("swal-input2").value
      ];
    }
  });
  if (formValues) {
    Swal.fire(JSON.stringify(formValues));
  }
}