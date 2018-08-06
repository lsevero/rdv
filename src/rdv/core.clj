(ns rdv.core
  (:require [dk.ative.docjure.spreadsheet :as ss]
            [clojure.java.io :as io]
            [seesaw.core :as ui]
            [seesaw.chooser :as ui-file]
            [clojure.data.csv :as csv]
            [clojure.string :as s])
  (:gen-class))

(defn in?
  "true if coll contains elm"
  [elm coll]
  (some #(= elm %) coll))

(defn today []
  (.format (java.text.SimpleDateFormat. "dd-MM-yyyy") (java.util.Date.)))

(def cell-name "B10")
(def cell-cpf "F10")
(def cell-departament "B12")
(def cell-op "G12")
(def cell-bank "C14")
(def cell-agency "E14")
(def cell-account "G14")
(def cell-reason "B18")
(def entries-per-sheet 20)
(def line-begin 4)
(def line-end 24)
(def rdv_base "rdv_base.xlsx")
(def rdv_amex )
(def column-date "A")
(def column-breakfast "C")
(def column-lunch "D")
(def column-dinner "E")
(def column-others "H")
(def column-reason "J")
(def column-hotel "F")
(def column-car "G")
(def code-breakfast ["CFDM" "BRFT"])
(def code-lunch ["ALM" "LNCH"])
(def code-dinner ["JANT" "DINN"])
(def code-hotel ["HTL"])
(def code-car ["ALGV" "RCAR"])
(def code-others ["OUT" "MISC"])
(def csv-code-column 5)
(def csv-price-column 2)
(def csv-reason-column 1)
(def csv-date-column 0)

;;; THE STATE OF THE APP
(def csv-file (atom ""))

(def pick-file (ui/button :text "Csv..."))

(def generate (ui/button :text "GERAR!!!" :enabled? true))

(defn extrai-info [dict]
  (conj {:csv @csv-file} dict))

(defn str->price [price-str]
  (Float/parseFloat (s/replace price-str "," ".")))

(defn set-breakfast! [n price sheet]
  (ss/set-cell! (ss/select-cell (str column-breakfast n) sheet) (str->price price)))

(defn set-lunch! [n price sheet]
  (ss/set-cell! (ss/select-cell (str column-lunch n) sheet) (str->price price)))

(defn set-dinner! [n price sheet]
  (ss/set-cell! (ss/select-cell (str column-dinner n) sheet) (str->price price)))

(defn set-hotel! [n price sheet]
  (ss/set-cell! (ss/select-cell (str column-hotel n) sheet) (str->price price)))

(defn set-car! [n price sheet]
  (ss/set-cell! (ss/select-cell (str column-car n) sheet) (str->price price)))

(defn set-others! [n price reason sheet]
  (ss/set-cell! (ss/select-cell (str column-others n) sheet) (str->price price))
  (ss/set-cell! (ss/select-cell (str column-reason n) sheet) reason))

(defn get-code [seq_]
  (nth seq_ csv-code-column))

(defn populate-page [i xls csv]
  (let [verso (ss/select-sheet (str "verso" i) xls)
        page (nth csv i)]
    (doseq [line-page (range (count page))]
      (let [line-xls (+ line-page line-begin)
            line-page-seq (nth page line-page)]
        (ss/set-cell! (ss/select-cell (str column-date line-xls) verso) (nth line-page-seq csv-date-column))
        (cond
          (in? (get-code line-page-seq) code-breakfast) (set-breakfast! line-xls (nth line-page-seq csv-price-column) verso)
          (in? (get-code line-page-seq) code-lunch) (set-lunch! line-xls (nth line-page-seq csv-price-column) verso)
          (in? (get-code line-page-seq) code-dinner) (set-dinner! line-xls (nth line-page-seq csv-price-column) verso)
          (in? (get-code line-page-seq) code-hotel) (set-hotel! line-xls (nth line-page-seq csv-price-column) verso)
          (in? (get-code line-page-seq) code-car) (set-car! line-xls (nth line-page-seq csv-price-column) verso)
          :else (set-others! line-xls (nth line-page-seq csv-price-column) (nth line-page-seq csv-reason-column) verso))))))

(defn populate-xls [dict]
  (let [xls (ss/load-workbook-from-resource rdv_base)
        csv (partition entries-per-sheet entries-per-sheet nil
                       (drop 1 (with-open [reader (io/reader (:csv dict))] (doall (csv/read-csv reader)))))
        capa (ss/select-sheet "capa" xls)]
    (ss/set-cell! (ss/select-cell cell-name capa) (:name dict))
    (ss/set-cell! (ss/select-cell cell-cpf capa) (:cpf dict))
    (ss/set-cell! (ss/select-cell cell-departament capa) (:departament dict))
    (ss/set-cell! (ss/select-cell cell-op capa) (:op dict))
    (ss/set-cell! (ss/select-cell cell-bank capa) (:bank dict))
    (ss/set-cell! (ss/select-cell cell-agency capa) (:agency dict))
    (ss/set-cell! (ss/select-cell cell-account capa) (:account dict))
    (ss/set-cell! (ss/select-cell cell-reason capa) (:reason dict))
    (doseq [i (range (count csv))]
      (populate-page i xls csv))
    (ss/save-workbook! (str (:name-saved-file dict) (today) ".xlsx") xls)))

(defn amex [dict]
  (populate-xls dict))

(defn gera-relatorios [dict]
  (if (empty? (:csv dict))
               (do
                 (ui/alert "Vacilão, escolhe um csv")
                 false)
               (do
                 (ui/config! generate :enabled? false :text "Gerando tabelas...")
                 (amex dict)
                 (ui/alert "Tabelas salvas!\nTa me devendo um pf")
                 (ui/config! generate :enabled? true :text "GERAR!!!")
                 true)))

(def form (ui/grid-panel :columns 2
                         :items ["Nome:" (ui/text :id :name)
                                 "CPF:" (ui/text :id :cpf)
                                 "Departamento:" (ui/text :id :departament)
                                 "OP:" (ui/text :id :op)
                                 "Banco:" (ui/text :id :bank)
                                 "Agência:" (ui/text :id :agency)
                                 "Conta corrente:" (ui/text :id :account)
                                 "Finalidade da viagem:" (ui/text :id :reason)
                                 "Tabela csv:" pick-file
                                 "Nome da nova tabela:" (ui/text :id :name-saved-file)
                                 "" generate]))

(ui/listen generate :mouse-clicked (fn [e]
                                     (future
                                       (->>
                                         form
                                         ui/value
                                         extrai-info
                                         gera-relatorios))))

(ui/listen pick-file :mouse-clicked (fn [e]
                                      (reset! csv-file (.getAbsolutePath
                                                       (ui-file/choose-file :type :open
                                                                            :multi? false
                                                                            :remember-directory? true
                                                                            :filters [["CSV files" ["csv"]]])))
                                      (ui/config! pick-file :text @csv-file)))
(defn -main
  "RAIO BARNABENIZADOR"
  [& args]
  (println "RAIO BARNABENIZADOR")
  (ui/invoke-later
    (-> (ui/frame :title "Barnabator"
                  :content form
                  :on-close :exit)
        ui/pack!
        ui/show!)))
