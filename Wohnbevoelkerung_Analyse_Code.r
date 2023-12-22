# Libraries laden
library(tidyverse)
library(readxl)
library(rstatix)

# Daten einlesen und verarbeiten
read_and_process_data <- function(path, sheet, year, range) {
  read_excel(path, sheet, range) %>%
    mutate(Jahr = year) %>%
    mutate(Kanton = str_extract(Region, "(?<=- ).*"),
           Bezirk = str_extract(Region, "(?<=>> ).*")) %>%
    tidyr::fill(c(Kanton, Bezirk), .direction = "down") %>%
    mutate(Kanton = ifelse(str_detect(Region, "^S"), NA, Kanton),
           Bezirk = ifelse(str_detect(Region, "^[S-]"), NA, Bezirk)) %>%
    separate(Region, into = c("Vorzeichen", "Region"), sep = " ", extra = "merge", fill = "left") %>%
    select(-Vorzeichen) %>%
    select(Kanton, Bezirk, Region, Jahr, Total, everything()) %>%
    rename(`100` = `100 und mehr`) %>%
    pivot_longer(`0`:`100`, names_to = "Alter", values_to = "Personen") %>%
    mutate(Alter = as.numeric(Alter),
           Alterskategorie = case_when(Alter < 18 ~ "Minderjährig",
                                      Alter < 65 ~ "Erwachsene jünger 65",
                                      .default =  "Erwachsene älter/gleich 65")) %>%
    group_by(Region, Kanton, Bezirk, Jahr, Alter, Alterskategorie) %>%
    summarise(Personen = sum(Personen, na.rm = TRUE)) %>%
    na.omit()
}

# Daten einlesen
daten10 <- read_and_process_data("su-d-01.02.03.06.xlsx", "2010", 2010, "A2:CY2763")
daten22 <- read_and_process_data("su-d-01.02.03.06.xlsx", "2022", 2022, "A2:CY2317")

# Daten zusammenfügen
daten_all <- bind_rows(daten10, daten22)

# Statistische Kennwerte berechnen
stat_kennwerte <- daten_all %>%
  group_by(Jahr, Alterskategorie) %>%
  summarise(across(Personen, list(n = ~n(), mn = mean, sd = ~sd(.), se = ~sd(.) / sqrt(n()),
                                  min = min, max = max, q1 = ~quantile(., 0.25),
                                  md = median, q3 = ~quantile(., 0.75),
                                  mad = mad, iqr = IQR)))

# Ergebnisse anzeigen und speichern
print(stat_kennwerte)
write.csv(stat_kennwerte, "statistische_kennwerte.csv", row.names = FALSE)

# Funktion zur Berechnung der Differenz des Durchschnittsalters
berechne_durchschnittsalter_diff <- function(data, kanton) {
  data %>%
    mutate(Alter = Alter + 1, altersAnzahl = Alter * Personen) %>%
    filter(Kanton == kanton & !is.na(Bezirk)) %>%
    group_by(Alterskategorie, Bezirk, Jahr) %>%
    summarise(durchschnittsalter = sum(altersAnzahl) / sum(Personen) - 1) %>%
    pivot_wider(names_from = Jahr, values_from = durchschnittsalter) %>%
    rename(Jahr10 = `2010`, Jahr22 = `2022`) %>%
    mutate(differenz_durchschnittsalter = Jahr22 - Jahr10)
}

# Berechnung für den Kanton Zürich
diff_durchschnittsalter_zh <- berechne_durchschnittsalter_diff(daten_all, "Zürich")

# Speichern der Differenzen als CSV-Datei
write.csv(diff_durchschnittsalter_zh, "diff_durchschnittsalter_zh.csv", row.names = FALSE)

# Ergebnis anzeigen
View(diff_durchschnittsalter_zh)


# Funktion zur Erstellung eines Boxplots mit spezifischem Design und Farben
create_custom_boxplot <- function(data, year_values) {
  # Definierte Farben für die Jahre
  year_colors <- c("2010" = "steelblue", "2022" = "darkorange")

  # Erstellen des Boxplots
  boxplot <- data %>%
    filter(Kanton == "Zürich", Alterskategorie == "Minderjährig", Bezirk != "Bezirk Zürich") %>%
    uncount(Personen) %>%
    filter(Jahr %in% year_values) %>%
    ggplot(aes(x = factor(Jahr), y = Alter, fill = factor(Jahr))) +
    geom_boxplot(outlier.shape = NA) +
    stat_summary(fun = median, geom = "crossbar", width = 0.75, fill = "white", color = "black", size = 0.5) +
    facet_wrap(~Bezirk) +
    labs(title = "Boxplot für Minderjährige in Zürich nach Bezirk", x = "Jahr", y = "Alter", fill = "Jahr") +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, vjust = 0.5, hjust = 1)) +
    scale_fill_manual(values = year_colors) +
    scale_y_continuous(breaks = seq(min(data$Alter), max(data$Alter), by = 1))

  return(boxplot)
}

# Boxplot mit spezifischem Design und Farben erstellen und anzeigen
boxplot_zh_custom <- create_custom_boxplot(daten_all, c(2010, 2022))
print(boxplot_zh_custom)

# Boxplot als PNG-Datei speichern
ggsave("boxplot_bezirke_zh_custom.png", plot = boxplot_zh_custom, width = 10, height = 8, dpi = 300)

# Weitere Aufgaben (Aufgabe 5, Aufgabe 6, etc.) können ähnlich angepasst werden.

# Einzelne Aufgaben anzeigen
View(daten_all)
View(daten22) # Beispiel, um die verarbeiteten Daten für 2022 anzuzeigen
