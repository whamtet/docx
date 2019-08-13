(ns ^{:doc "We wish to replace [Template String] with custom text.
      The catch is that [Template String] may be split accross several xml elements in the docx format.
      This namespace attempts to rearrange each [Template String] to fit inside a single xml element."}
  docx.cleanup
  (:require
    [clojure.string :as string]
    [clojure.walk :as walk]))

(defn split-after
  "splits s after f returns true"
  [f s]
  (let [[a b] (split-with (complement f) s)]
    [(concat a (take 1 b)) (rest b)]))

(defn break-with
  "starts a new subseqence every time pred returns true.
  The first subsequence contains elements with pred false"
  [pred s]
  (let [pred #(-> % pred boolean) ;force boolean
        broken (partition-by pred s)
        rejoin (fn [[a b]] (concat a b))]
    (if (-> s first pred)
      (cons [] (map rejoin (partition-all 2 broken)))
      (cons (first broken) (map rejoin (partition-all 2 (rest broken)))))))

(defn combine-seq [break take combine s]
  (let [[head & rest] (break-with break s)]
    (concat
      head
      (mapcat
        (fn [section]
          (let [[a b] (split-after take section)]
            (concat (combine a) b)))
        rest))))

(defn r-text [r]
  (when
    (and
      (vector? r)
      (-> r first (= :w:r))
      (-> r last first (= :w:t)))
    (-> r last last)))

(defn score [s]
  (reduce
    (fn [s c]
      (case c \[ (inc s) \] (dec s) s))
    0
    s))

(defn assoc-last [v s]
  (assoc v (dec (count v)) s))

(defn combine-text [elements]
  (let [{other-elements false text-elements true} (group-by #(boolean (r-text %)) elements)
        last-t (-> text-elements last last)
        combined-text (string/join "" (map r-text text-elements))
        new-t (assoc-last last-t combined-text)]
    (cons
      (assoc-last (first elements) new-t)
      other-elements)))

(defn merge-elements [elements]
  (combine-seq
    #(some-> % r-text score pos?)
    #(some-> % r-text score neg?)
    combine-text
    elements))

(defn merge-p [p]
  (if (and (vector? p) (= :w:p (first p)))
    (vec (merge-elements p))
    p))

(defn cleanup [doc]
  (walk/prewalk merge-p doc))
