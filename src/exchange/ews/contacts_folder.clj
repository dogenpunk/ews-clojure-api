(ns exchange.ews.contacts-folder
  (:require [exchange.ews.authentication :refer [service-instance]]
            [exchange.ews.folder :as ews-folder])
  (:import (microsoft.exchange.webservices.data.core.service.folder.Folder ContactsFolder)
           (microsoft.exchange.webservices.data.core.enumeration.property WellKnownFolderName)
           (microsoft.exchange.webservices.data.property.complex ItemId)
           (microsoft.exchange.webservices.data.search ItemView)
           (microsoft.exchange.webservices.data.search.filter SearchFilter$IsEqualTo)

           (microsoft.exchange.webservices.data.core.service.schema FolderSchema)))

(defn get-folder
  []
  (ews-folder/get-folder WellKnownFolderName/Contacts))

(defn get-sub-folder
  ([name]
   (let [sf (SearchFilter$IsEqualTo. FolderSchema/DisplayName name)
         items (.findItems @service-instance WellKnownFolderName/Contacts sf (ItemView. 1))]
     (first (.getItems items)))))
