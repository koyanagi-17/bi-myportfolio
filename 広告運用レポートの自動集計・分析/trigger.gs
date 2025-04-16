function trigger(){
  date = new Date();
  var month = date.getMonth();

  date.setDate(date.getDate()+7);
  var month_next_week = date.getMonth();

  if (month != month_next_week){
    multiprocessor_1();
  }
}
