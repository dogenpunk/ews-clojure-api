(ns exchange.ews.contact-group
  (:require [exchange.ews.authentication :refer [service-instance]]
            [clojure.set])
  (:import (microsoft.exchange.webservices.data.core.service.item ContactGroup)
           (microsoft.exchange.webservices.data.search ItemView)
           (microsoft.exchange.webservices.data.search.filter SearchFilter$IsEqualTo)
           (microsoft.exchange.webservices.data.core.service.schema ContactGroupSchema)
           (microsoft.exchange.webservices.data.core.enumeration.property WellKnownFolderName)
           (microsoft.exchange.webservices.data.core.enumeration.property EmailAddressKey)
           (microsoft.exchange.webservices.data.property.complex GroupMember GroupMemberCollection)
           (microsoft.exchange.webservices.data.core.enumeration.service ConflictResolutionMode)
           (microsoft.exchange.webservices.data.core.enumeration.property BasePropertySet)
           (microsoft.exchange.webservices.data.core PropertySet)
           (microsoft.exchange.webservices.data.property.complex GroupMemberCollection
                                                                 GroupMember
                                                                 ItemId)))


(defmulti get-contact-group
  type)

(defmethod get-contact-group java.lang.String
  [id]
  (ContactGroup/bind @service-instance (ItemId/getItemIdFromString id)))

(defmethod get-contact-group ItemId
  [id]
  (ContactGroup/bind @service-instance id (PropertySet. BasePropertySet/FirstClassProperties)))

(defn get-contact-group-by-name
  ([group-name]
     (let [sf (SearchFilter$IsEqualTo. ContactGroupSchema/DisplayName group-name)
           items (.findItems @service-instance WellKnownFolderName/Contacts sf (ItemView. 1))
           group (first (.getItems items))]
       (ContactGroup/bind @service-instance (.getId group) (PropertySet. BasePropertySet/FirstClassProperties)))))

(defn set-group-display-name
  [id display-name]
  (let [group (get-contact-group-by-id id)]
    (.setDisplayName group display-name)
    (.update group ConflictResolutionMode/AlwaysOverwrite)))

(defn create-group!
  ([display-name]
   (let [group (ContactGroup. @service-instance)]
     (.setDisplayName group display-name)
     (.save group)
     group))
  ([display-name members]
   (let [group (ContactGroup. @service-instance)
         group-members (.getMembers group)]
     (.setDisplayName group display-name)
     (doseq [{:keys [id emailAddress]} members]
       (.addPersonalContact group-members id emailAddress))
     (.save group)
     group)))

(defn get-contact-group-by-name
  ([group-name]
   (let [sf (SearchFilter$IsEqualTo. ContactGroupSchema/DisplayName group-name)
         items (.findItems @service-instance WellKnownFolderName/Contacts sf (ItemView. 1))]
     (first (.getItems items)))))

(defn get-contact-group-members
  [group-name]
  (let [cgroupId (.getId (get-contact-group-by-name group-name))
        cgroup (ContactGroup/bind @service-instance cgroupId)]
    (into [] (.getMembers cgroup))))


(defn add-contact!
  [group contact]
  (.addPersonalContact (.getMembers group) (.getId contact))
  group)

(defn remove-member!
  [group member]
  (let [members (.getMembers group)]
    (.remove members member)))

(defn get-members
  ([group]
   (.getMembers group))
  ([group emails]
   {:pre [(set? emails)]}
   (for [member (.getMembers group)
         :let [address (.getAddress (.getAddressInformation member EmailAddressKey/EmailAddress1))]
         :when (emails address)]
     member)))

(defn update-group!
  [group]
  (.update group ConflictResolutionMode/AlwaysOverwrite)
  group)
