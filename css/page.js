function test(){ 
          if (document.form2.page.value=="" || isNaN(document.form2.page.value))  {
               alert("页码必须为阿拉伯数字，请重新输入");
               document.form2.page.focus();
               return false;
              }
  return true;
}
