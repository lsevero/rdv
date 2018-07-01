(defproject rdv "0.1.0"
  :description "RDV"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.8.0"]
                 [seesaw "1.5.0"]
                 [dk.ative/docjure "1.12.0"]
                 [org.clojure/data.csv "0.1.4"]]
  :main ^:skip-aot rdv.core
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all}})
