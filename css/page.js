function test(){ 
          if (document.form2.page.value=="" || isNaN(document.form2.page.value))  {
               alert("ҳ�����Ϊ���������֣�����������");
               document.form2.page.focus();
               return false;
              }
  return true;
}
