(ns docx.file
  (:refer-clojure :exclude [read])
  (:require
    [korum-platform.config :refer [dev?]]
    [docx.xml :as xml]
    [docx.cleanup :as cleanup]
    [clojure.contrib.prxml :refer [prxml]]
    [clojure.java.io :as io])
  (:import
    (java.nio.file
      Paths
      FileSystems
      Files
      OpenOption
      )
    org.apache.commons.io.FileUtils
    java.io.File
    (java.nio.file.attribute
      FileAttribute)))

(def s (make-array String 0))
(def o (make-array OpenOption 0))
(def a (make-array FileAttribute 0))

(defn temp-file []
  (str (Files/createTempFile nil nil a)))

(defn ->readable [{:keys [tag attrs content] :as contents}]
  (if tag
    (vec
      (list* tag attrs (map ->readable content)))
    contents))

(defn slurp-xml [f]
  (with-open [in (Files/newInputStream f o)]
    (-> in xml/parse ->readable)))

(defn spit-xml [f xml]
  (with-open [out (Files/newOutputStream f o)]
    (-> xml prxml with-out-str (io/copy out))))

(defn raw [f]
  (FileUtils/readFileToByteArray f))

(use 'clojure.pprint)
(.mkdir (File. "tmp"))
(defn spit-edn [i s]
  (spit (format "tmp/edit%s.edn" i)
        (with-out-str (pprint s))))

(defn edit
  ([f func]
   (let [temp-file (temp-file)
         _ (edit (str f) func temp-file)
         data (FileUtils/readFileToByteArray (File. temp-file))]
     (.delete (File. temp-file))
     data))
  ([f1 func f2]
   (io/copy (File. f1) (File. f2))
   (with-open [fs1 (FileSystems/newFileSystem (Paths/get f1 s) nil)
               fs2 (FileSystems/newFileSystem (Paths/get f2 s) nil)]
     (let [doc1 (.getPath fs1 "/word/document.xml" s)
           doc2 (.getPath fs2 "/word/document.xml" s)
           xml1 (slurp-xml doc1)
           xml2 (-> xml1 cleanup/cleanup func)]
       (when false ;(dev?)
         (io/copy (Files/newInputStream doc1 o) (File. "tmp/edit0.xml"))
         (spit-edn 1 xml1)
         (spit-edn 2 (cleanup/cleanup xml1))
         (spit-edn 3 xml2)
         (spit "tmp/edit4.xml" (-> xml2 prxml with-out-str)))
       (spit-xml doc2 (or xml2 xml1))))))

(defn read [f func]
  (with-open [fs (FileSystems/newFileSystem (Paths/get f s) nil)]
    (-> fs
        (.getPath "/word/document.xml" s)
        slurp-xml
        cleanup/cleanup
        func)))

(defn test-file []
  (edit
    "contracts/UK (Company) KORUM CONSULTANCY T&C (Aug 2019)/UK (Company) KORUM CONSULTANCY T&C (Aug 2019).docx"
    (fn [_]
      (-> "tmp/edit2.edn" slurp read-string))
    "/Users/matthew/Downloads/test.docx"))
