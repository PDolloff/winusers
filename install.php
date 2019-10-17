<?php

/**
 * This function is called on installation and is used to create database schema for the plugin
 */
function extension_install_winusers()
{
    $commonObject = new ExtensionCommon;
    $commonObject -> sqlQuery("CREATE TABLE IF NOT EXISTS `winusers` (
                             `ID` INT(11) NOT NULL AUTO_INCREMENT,
                             `HARDWARE_ID` INT(11) NOT NULL,
                             `NAME` VARCHAR(255) DEFAULT NULL,
                             `LOGINTIME` datetime DEFAULT NULL,
                             PRIMARY KEY  (`ID`,`HARDWARE_ID`)
                             ) ENGINE=InnoDB ;");
}
/**
 * This function is called on removal and is used to destroy database schema for the plugin
 */
function extension_delete_winusers()
{
    $commonObject = new ExtensionCommon;
    $commonObject -> sqlQuery("DROP TABLE `winusers`;");
}

/**
 * This function is called on plugin upgrade
 */
function extension_upgrade_winusers()
{
}
