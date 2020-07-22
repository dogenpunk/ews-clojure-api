(ns exchange.ews.contact
  (:require [exchange.ews.authentication :refer [service-instance]])
  (:import (microsoft.exchange.webservices.data.core.service.item Contact)
           (microsoft.exchange.webservices.data.core.enumeration.service ConflictResolutionMode)
           (microsoft.exchange.webservices.data.property.complex CompleteName
                                                                 EmailAddress
                                                                 EmailAddressCollection
                                                                 EmailAddressDictionary
                                                                 ImAddressDictionary
                                                                 ItemAttachment
                                                                 ItemId
                                                                 PhoneNumberDictionary
                                                                 PhysicalAddress
                                                                 StringList)))

(defn get-contact
  [id]
  (Contact/bind @service-instance (ItemId/getItemIdFromString id)))

(defn set-contact-display-name
  [id display-name]
  (let [contact (get-contact id)]
    (.setDisplayName contact display-name)
    (.update contact ConflictResolutionMode/AutoResolve)))

(defn set-contact-given-name
  [id given-name]
  (let [contact (get-contact id)]
    (.setGivenName contact given-name)
    (.update contact ConflictResolutionMode/AutoResolve)))

(defn set-contact-surname
  [id surname]
  (let [contact (get-contact id)]
    (.setSurname contact surname)
    (.update contact ConflictResolutionMode/AutoResolve)))
