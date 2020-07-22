(ns exchange.ews.contact-group
  (:require [exchange.ews.authentication :refer [service-instance]])
  (:import (microsoft.exchange.webservices.data.core.service.item ContactGroup)
           (microsoft.exchange.webservices.data.core.enumeration.service ConflictResolutionMode)
           (microsoft.exchange.webservices.data.property.complex GroupMemberCollection
                                                                 GroupMember
                                                                 ItemId)))

(defn get-contact-group
  [id]
  (ContactGroup/bind @service-instance (ItemId/getItemIdFromString id)))

(defn set-group-display-name
  [id display-name]
  (let [group (get-contact-group id)]
    (.setDisplayName group display-name)
    (.update group ConflictResolutionMode/AutoResolve)))

(defn create-group
  ([display-name]
   (let [group (ContactGroup. @service-instance)]
     (.setDisplayName group display-name)
     (.save group)))
  ([display-name members]
   (let [group (ContactGroup. @service-instance)
         group-members (.getMembers group)]
     (.setDisplayName group display-name)
     (doseq [{:keys [id emailAddress]} members]
       (.addPersonalContact id emailAddress))
     (.save group))))

(comment
  (create-group
   "Instructors - Inactive"))
