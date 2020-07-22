(defproject com.temify/ews-clojure-api "0.0.9"
  :description "Utility library for accessing Microsoft Exchange"
  :url "https://wwww.bizziapp.com"

  :plugins [[lein-codox "0.10.4"]
            [lein-tools-deps "0.4.5"]]
  :middleware [lein-tools-deps.plugin/resolve-dependencies-with-deps-edn]
  :lein-tools-deps/config {:config-files [:install :user :project]}
  :java-source-paths ["src/java"])
