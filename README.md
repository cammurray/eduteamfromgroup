# EDU Team from Group

## Background

The following work was performed as an interim solution whilst a university could not deploy School Data Sync (SDS). The university had classes being created by the class management system in Azure AD as standard groups. The purpose of this script was to, without changing the class management system, have Teams created for the respective classes created by the class management system.

The script supports the concept of a 'class' - created using the EDU template, and a 'program' created using a standard teams template.

## School Data Sync (SDS)

School Data Sync https://sds.microsoft.com/ should preferentially be used over top of any custom solution such as this one.

## Customisations

The script will require customisations in order to work in your tenant. Most of these customisations are explained in the parameter block at the start of the script.

## Requirements

1. Customise the script to your requirements, it won't work 'out of the box'. Use this as a template for creating your own script.
2. Azure Automation - reality is it will probably work outside of Azure Automation, but some changes will need to be made.
3. An Azure AD Application
4. A username for creating the groups - some modifications could be made in order to support application secrets instead of a username/password
5. A group to use as the 'master group', only groups added to this group will have a respective team created for them.

## Warranty

This script will not work out of the box - you'll need to customise it. Use it as a template and a starting point for your own script. It's being shared to the wider community incase there is a similar use case only.