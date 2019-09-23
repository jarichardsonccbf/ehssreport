example_pp <- read_pptx() %>% 
  add_slide(layout = "Title Slide", master = "Office Theme") %>% 
  ph_with_text(
    type = "ctrTitle",
    str = "Weekly P3 Deck"
  ) %>% 
  ph_with(
    location = ph_location_type(type = "subTitle"),
    value = "Copy and paste the generated tables into your report"
  ) %>% 
  
  # LMS slide ----
# paste("EHSS - Compliance Training Completion Status", months(Sys.Date() - months(1)), "/", months(Sys.Date()), year(Sys.Date()), "as of", format(Sys.Date(), format ="%m/%d/%Y"))

add_slide(layout = "Title and Content", master = "Office Theme") %>% 
  ph_with(
    block_list(
      fpar(fp_p = fp_par(text.align = "center"),
        ftext(
          "test", 
          prop = fp_text(font.size = 28)
             )
          )
              ),
          location = ph_location_type(type = "title")
        )

example_pp

print(example_pp, target = "filehsstest.pptx")
