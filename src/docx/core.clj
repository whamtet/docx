(ns docx.core
  (:require
    [docx.file :as file]
    [clojure.string :as string]
    [clojure.walk :as walk]))

(defn replace-all [s m]
  (reduce
    (fn [s [k v]] (.replace s k v))
    s
    m))

(defn walk-doc [doc m]
  (walk/prewalk #(if (string? %) (replace-all % m) %) doc))
(defn edit-doc [f m]
  (let [m (into {} (for [[k v] m
                         :when (and v (not (#{"null" "undefined"} v)))]
                     [(str "[" k "]") v]))]
    (file/edit f #(walk-doc % m))))

(defn get-strs [coll]
  (cond
    (and (vector? coll) (= :w:t (first coll))) [(last coll)]
    (coll? coll) (mapcat get-strs coll)))

(defn doc-strs [doc]
  (string/join "" (get-strs doc)))

(defn keywords [doc]
  (->> doc
       doc-strs
       (re-seq #"\[([^\]]+)")
       (map second)
       (remove #(.contains % "req|signer"))
       distinct))

(defn doc-keywords [f]
  (file/read f keywords))

(def raw file/raw)
