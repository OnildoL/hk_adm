class Fields {
  page := ""

  __new(page) {
    this.page := page
  }

  GetsTheValueOfAnIntegerFromaField(xpath) {
    repeatGetsTheValueOfAnIntegerFromaField:
    result := this.page.getElementsByXpath(xpath)[0].value

    if result is integer 
    {
      return result
    } else {
      goto, repeatGetsTheValueOfAnIntegerFromaField
    }
  }
}