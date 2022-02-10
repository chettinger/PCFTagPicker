import * as React from "react";

import {
    TagPicker,
    ITag,
    IBasePickerSuggestionsProps,
} from "@fluentui/react/lib/Pickers";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { useId } from "@fluentui/react-hooks";

const rootClass = mergeStyles({
    maxWidth: 500,
});

const pickerSuggestionsProps: IBasePickerSuggestionsProps = {
    suggestionsHeaderText: "Suggested tags",
    noResultsFoundText: "No tags found",
};
let tags: ITag[] = [];

const listContainsTagList = (tag: ITag, tagList?: ITag[]) => {
    if (!tagList || !tagList.length || tagList.length === 0) {
        return false;
    }
    return tagList.some((compareTag) => compareTag.key === tag.key);
};

const filterSuggestedTags = (filterText: string, tagList?: ITag[]): ITag[] => {
    return filterText
        ? tags.filter(
              (tag) =>
                  tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0 && !listContainsTagList(tag, tagList),
          )
        : [];
};

const getTextFromItem = (item: ITag) => item.name;
export interface IRecordSelectorProps {
    availableTags: ITag[];
    selectedTags: ITag[];
    onChange: (items?: ITag[]) => void;
}
export const ConnectionTagPicker: React.FunctionComponent<IRecordSelectorProps> = (props) => {
    const pickerId = useId("inline-picker");
    tags = props.availableTags;
    return (
        <div className={rootClass}>
            <TagPicker
                removeButtonAriaLabel="Remove"
                selectedItems={props.selectedTags}
                onResolveSuggestions={filterSuggestedTags}
                getTextFromItem={getTextFromItem}
                pickerSuggestionsProps={pickerSuggestionsProps}
                pickerCalloutProps={{ doNotLayer: true }}
                inputProps={{
                    id: pickerId,
                }}
                onChange={(items?: ITag[]) => {
                    props.onChange(items);
                }}
            />
            <div
                style={{ height: "10em" }}
            />
        </div>
    );
};
