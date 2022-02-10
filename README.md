# PCFTagPicker

This project utilizes the Fluent UI TagPicker control to render Connections to an entity record as tags.

## Requirements

* A tag category entity (ex. ```my_tagcategory```)
  * A **Single Line of Text** for the category name (ex. ```my_name```)
* A tag entity (ex. ```my_tag```)
  * A **Single Line of Text** for the tag name (ex. ```my_name```)
* A lookup to the category entity
* [Connections enabled](https://docs.microsoft.com/en-us/powerapps/maker/data-platform/configure-connection-roles#enable-connection-roles-for-a-table) for the tag entity and the entity to be tagged
* [Connection Role](https://docs.microsoft.com/en-us/powerapps/maker/data-platform/configure-connection-roles#add-connection-roles-to-a-solution) added with Tag entity as record type. Recommended:
  * Step 1: Name = Tag Connections
  * Step 2: Only these record types = Tag

## Control Configuration

The control can be used for **Single Line of Text** or **Multiple Lines of Text** field types. Even though the control manages connections between entities, it also stores a comma-delimited string in the field for display purposes. _**Warning:**_ Keep this in mind when choosing a field type. For example, if many tags are added to a record and a **Single Line of Text** field is used with a limit of 200 characters; you could encounter an error. It is recommended to use a **Multiple Lines of Text** field type.

The control requires several items to be configured:

* TagsField: The Field the control is bound to (ex. ```my_tags```)
* TagsEntity: The name of the Tag entity (ex. ```my_tag```)
* TagsEntityDisplayNameField: The Tag field to display in the control (ex. ```my_name```)
* TagsEntityCategoryField: The field of the tag entity that references the category (ex. ```my_category```)
* ConnectionRoleName: The name of the Role used for tags connection (ex. Tags). NOTE: Connections must be enabled on entities and Role created.
* CategoryEntity: The name of the Category entity (ex. ```my_tagcategory```)
* CategoryEntityNameField: The Category name field (ex. ```my_name```)
* CategoriesField (optional): The Categories to filter (comma-delimited) (ex. Default, Category 1,Category 2). If nothing in provided, then all tags will be returned. Otherwise, enter the category names to filter.

## Searching

In order to search for entity records that have connected tags, do the following:

### Using Connections

This approach uses the actual tags to filter results

1. Open Advanced Search
1. Select the entity with connected tags
1. Select **Connection (Connected From)**
1. Select **Connected To** for field
1. Select **Equals** for operator
1. Using the lookup tool:
    * Look for: Select Tag entity
    * Select the tags for your criteria and then click **Add**
1. Configure columns and any additional criteria and then click **Results**

### Using String

This approach uses the comma-delimited string stored on the record to filter results

1. Open Advanced Search
1. Select the entity with connected tags
1. Select **Tags** for field
1. Select **Contains** for operator
1. Enter the text of the tag you are looking for
1. Repeat previous 3 steps with grouping if desired to refine search criteria
1. Configure columns and any additional criteria and then click **Results**
