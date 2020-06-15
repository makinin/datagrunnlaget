## Laster pakker ----
library(rdbhapi)
library(tidyverse)
library(openxlsx)

## Leser og filtrerer data til finansieringssystemet ----

# DBH-tabeller med API-valg

finsys_dbh_spørringer <-
  tribble(~tabellnavn, ~table_id, ~group_by, ~filters,
          "studiepoeng", 900, NULL, list(Årstall = c("top", 2)),
          "kandidater", 907, NULL, list(Årstall = c("top", 2)),
          "utveksling", 142, c("Årstall", "Institusjonskode", "Utvekslingsavtale","Type", "Nivåkode"), list(Årstall = c("top", 2)),
          "økonomi", 902, NULL, list(Årstall = c("top", 2)),
          "publisering", 373, c("Årstall", "Institusjonskode"), list("Årstall" = c("top", 2), "Kode for type publiseringskanal"=c("1","2")),
          "doktorgrader", 101, c("Institusjonskode", "Årstall"), list(Årstall = c("top", 2)),
          "doktorgrader_samarbeid", 100, c("Årstall", "Institusjonskode (arbeidsgiver)"), list(Årstall = c("top", 2)),
          "PKU", 98, NULL, list(Årstall = c("top", 2)),
          "institusjoner", 211, NULL, NULL,
  )


# Laster ned alle finsystabellene

finsys_dbh <- 
  local({
    pb <- progress_estimated(nrow(finsys_dbh_spørringer))
    res <- 
      pmap(finsys_dbh_spørringer,
           function(tabellnavn, ...) {
             pb$tick()$print()
             res <- 
               do.call(dbh_data,
                       list(...)) %>% 
               rename_all(str_to_lower)
             res <- setNames(list(res), tabellnavn)
             res
           })
    pb$stop()$print()
    unlist(res, recursive = FALSE)
  })

finsys <- finsys_dbh

# Leser oversikt over institusjoner som er inkludert i finansieringssystemet, og unntatt fra enkeltindikatorer

finsys_utvalg <- 
  c("institusjoner", "unntak") %>%
  setNames(.,.) %>% 
  map(~read_tsv(file.path("data", sprintf("finsys_%s.tsv", .)),
                col_types = cols(budsjettår = col_integer(),
                                .default = col_character())))

# Viderefører inneværende utvalg
finsys_utvalg$institusjoner <- 
  finsys_utvalg$institusjoner %>% 
  complete(institusjonskode, budsjettår = 2020:2021)

finsys_utvalg$unntak <- 
  finsys_utvalg$unntak %>% 
  complete(nesting(institusjonskode, indikator), budsjettår = 2020:2021)

## Filtrerer finsystabeller til hva som gir uttelling, og harmoniserer
## variabelnavn og -verdier til sammenslåing

finsys$studiepoeng <-
  finsys_dbh$studiepoeng %>%
  rename(indikatorverdi = `ny produksjon egentfin`) %>%
  # Bruker finansieringskategorien basert på studentens tilhørighet for BI og for emnet ellers
  mutate(kategori =
           case_when(institusjonskode == "8241" ~
                       `finmodkode student`,
                     TRUE ~
                       `finmodekode emne`)) %>% 
  mutate_at(c("studentkategori", "kategori"), str_to_upper) %>% 
  filter(kategori %in% LETTERS[1:6],
         studentkategori == "S")

finsys$kandidater <-
  finsys_dbh$kandidater %>% 
  rename(faktor = `uttelling kode:1=enkel, 2=dobbel`,
         etterrapportert = etterrapp,
         innpasset = `dobbel til enkel`,
         kategori = `finansierings-kategori`) %>% 
  mutate_at("kategori", str_to_upper) %>%
  mutate_at("faktor", parse_integer) %>%
  mutate_at(c("etterrapportert", "innpasset"), replace_na, 0) %>% 
  mutate(ordinær = totalt - etterrapportert - innpasset) %>% 
  select(-totalt) %>% 
  gather(ordinær, etterrapportert, innpasset,
         key = "kandidatgruppe",
         value = "indikatorverdi")

finsys$erasmus_plus<-finsys_dbh$utveksling_org %>% 
  rename(indikatorverdi = `antall totalt`) %>% 
  mutate_at(c("utvekslingsavtale", "type"), str_to_upper) %>%
  filter(utvekslingsavtale == "ERASMUS+" & nivåkode!= "FU" & type=="NORSK") %>% 
  mutate(kategori="Erasmus+")

finsys$utenlandsk<-finsys_dbh$ utveksling_org %>% 
  rename(indikatorverdi = `antall totalt`) %>% 
  mutate_at(c("utvekslingsavtale", "type"), str_to_upper) %>%
  filter(utvekslingsavtale != "INDIVID" & nivåkode!= "FU" & type=="UTENL") %>% 
  mutate(kategori="ordinær")

finsysutreisende_norsk<- finsys_dbh$utveksling_org %>% 
  rename(indikatorverdi = `antall totalt`) %>% 
  mutate_at(c("utvekslingsavtale", "type"), str_to_upper) %>%
  filter(!utvekslingsavtale %in% c("ERASMUS+","INDIVID") & nivåkode!= "FU" & type=="NORSK") %>% 
  mutate(kategori="ordinær")

finsys$utveksling<-bind_rows(finsys$erasmus_plus,finsys$utenlandsk,finsys$utreisende_norsk)



finsys$doktorgrader <- 
  list(ordinær = finsys_dbh$doktorgrader,
       `samarbeids-ph.d.` = finsys_dbh$doktorgrader_samarbeid,
       PKU = finsys_dbh$PKU) %>% 
  map_dfr(~setNames(., str_replace_all(names(.),
                                       c("antall($| totalt)" = "indikatorverdi",
                                         "institusjonskode.*" = "institusjonskode"))),
          .id = "kandidatgruppe") %>% 
  mutate(faktor = case_when(kandidatgruppe == "samarbeids-ph.d." ~ 0.2))
  
finsys$publisering <- 
  finsys_dbh$publisering %>% 
  rename(indikatorverdi = publiseringspoeng)

finsys$økonomi <- 
  finsys_dbh$økonomi %>% 
  rename(EU = eu) %>% 
  mutate(forskningsråd = map2_dbl(nfr, rff, sum, na.rm = TRUE),
         BOA = map2_dbl(bidragsinntekter, oppdragsinntekter, sum, na.rm = TRUE))

finsys <- 
  c("EU", "forskningsråd", "BOA") %>% 
  setNames(., .) %>% 
  map(~select_at(finsys$økonomi,
                 c("institusjonskode", "årstall", "indikatorverdi" = .))) %>% 
  c(finsys)

## Lager samletabell ----

finsys_data <-
  finsys[c("studiepoeng", "kandidater", "utveksling", "doktorgrader", "publisering", "EU", "forskningsråd", "BOA")] %>%
  bind_rows(.id = "indikator") %>%
  group_by(
    årstall,
    institusjonskode,
    indikator, 
    kategori, 
    kandidatgruppe,
    faktor) %>%
  summarise_at("indikatorverdi", sum, na.rm = TRUE) %>%
  ungroup %>% 
  mutate(budsjettår = årstall + 2L,
         indikatorverdi_produkt = indikatorverdi * replace_na(faktor, 1)) %>% 
  semi_join(finsys_utvalg$institusjoner,
            by = c("institusjonskode", "budsjettår")) %>% 
  anti_join(finsys_utvalg$unntak,
            by = c("institusjonskode", "budsjettår", "indikator"))

finsys$institusjoner <-
  finsys_dbh$institusjoner %>%
  rename(institusjonskode_nyeste = `institusjonskode (sammenslått)`,
         institusjonsnavn_nyeste = `sammenslått navn`) %>% 
  semi_join(finsys_data, by = "institusjonskode") %>% 
  left_join(select(finsys$institusjoner,
                   institusjonskode, kortnavn_nyeste = kortnavn),
            by = c("institusjonskode_nyeste" = "institusjonskode"))

## Eksporterer samletabell til Excel-format ----

lagre_excel <- 
  function(tabeller, filnavn) {
  wb <- createWorkbook()
  iwalk(tabeller,
       function(tabell, tabellnavn) {
         addWorksheet(wb, tabellnavn)
         writeDataTable(wb, tabellnavn, tabell,
                        tableStyle = "TableStyleLight1",
                        tableName = tabellnavn)
       })
  saveWorkbook(wb, filnavn, overwrite = TRUE)
  }

lagre_excel(list(finsys_data = finsys_data,
                 institusjoner = finsys$institusjoner),
            "finsys_data.xlsx")


## Lager tabeller med produksjonsdata, tilsvarende dem i Blått hefte ----

lag_produksjonstabell <- function(df,
                                  filter_indikatorer, # indikatorene som skal inngå i tabellen
                                  kolonne_vars, # variablene som skal spres på kolonner (indikator/kategori/osv.)
                                  funs # tilleggsfunksjon som skal anvendes på kolonnene (for å lage andelstall)
) {
  if (is.null(funs)) {
    funs <- identity
  }
  df %>% 
    filter(indikator %in% filter_indikatorer) %>% 
    unite_("kolonne_vars_samlet", kolonne_vars) %>% 
    group_by_at(c("institusjonskode_nyeste",
                  "institusjonsnavn_nyeste",
                  "kortnavn_nyeste",
                  "kolonne_vars_samlet",
                  "budsjettår")) %>% 
    summarise_at("indikatorverdi", sum) %>% 
    ungroup() %>% 
    spread(key = "kolonne_vars_samlet", value = "indikatorverdi") %>%
    mutate_at(vars(everything(), -matches("institusjon|navn|år")), funs)
}

produksjonstabeller <- 
  tribble(~filter_indikatorer, ~kolonne_vars, ~funs,
          "studiepoeng", "kategori", NULL,
          "kandidater", c("kategori", "faktor"), NULL,
          c("doktorgrader", "utveksling"), c("indikator", "faktor", "kategori"), NULL,
          c("publisering", "EU", "forskningsråd", "BOA"), c("indikator") , list(pst = function(x) 100 * x / sum(x, na.rm = TRUE))) %>% 
  pmap(function(...) {
    finsys_data %>% 
      filter(budsjettår == 2021) %>%
      left_join(select(finsys$institusjoner,
                       institusjonskode, institusjonskode_nyeste, institusjonsnavn_nyeste, kortnavn_nyeste),
                by = "institusjonskode") %>% 
      lag_produksjonstabell(...)
  }) 

names(produksjonstabeller) <-
  c("studiepoeng", "kandidater", "utveksling_doktorgrader", "lukket_ramme")

lagre_excel(produksjonstabeller,
            "produksjonstabeller.xlsx")
