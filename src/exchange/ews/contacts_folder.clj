(ns exchange.ews.contacts-folder
  (:require [exchange.ews.authentication :refer [service-instance]])
  (:import (microsoft.exchange.webservices.data.core.service.folder.Folder ContactsFolder)
           (microsoft.exchange.webservices.data.property.complex ItemId)))
