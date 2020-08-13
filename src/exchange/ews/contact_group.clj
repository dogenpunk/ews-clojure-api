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
  (ContactGroup/bind @service-instance id))

(defn get-contact-group-by-name
  ([group-name]
     (let [sf (SearchFilter$IsEqualTo. ContactGroupSchema/DisplayName group-name)
           items (.findItems @service-instance WellKnownFolderName/Contacts sf (ItemView. 1))]
       (first (.getItems items)))))

(defn set-group-display-name
  [id display-name]
  (let [group (get-contact-group id)]
    (.setDisplayName group display-name)
    (.update group ConflictResolutionMode/AlwaysOverwrite)))

(defn create-group!
  ([display-name]
   (let [group (ContactGroup. @service-instance)]
     (.setDisplayName group display-name)
     (.save group)
     group)))

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
