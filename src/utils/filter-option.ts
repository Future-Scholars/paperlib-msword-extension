import { stringUtils } from "paperlib-api/utils";

export interface IPaperFilterOptions {
  search?: string;
  searchMode?: "general" | "fulltext" | "advanced";
  flaged?: boolean;
  tag?: string;
  folder?: string;
  limit?: number;
}

export class PaperFilterOptions implements IPaperFilterOptions {
  public filters: string[] = [];
  public search?: string;
  public searchMode?: "general" | "fulltext" | "advanced";
  public flaged?: boolean;
  public tag?: string;
  public folder?: string;
  public limit?: number;

  constructor(options?: Partial<IPaperFilterOptions>) {
    if (options) {
      this.update(options);
    }
  }

  update(options: Partial<IPaperFilterOptions>) {
    for (const key in options) {
      this[key] = options[key];
    }

    this.filters = [];

    if (this.search) {
      let formatedSearch = stringUtils.formatString({
        str: this.search,
        removeNewline: true,
        trimWhite: true,
      });

      if (!this.searchMode || this.searchMode === "general") {
        const fuzzyFormatedSearch = `*${formatedSearch
          .trim()
          .split(" ")
          .join("*")}*`;
        this.filters.push(
          `(title LIKE[c] \"${fuzzyFormatedSearch}\" OR authors LIKE[c] \"${fuzzyFormatedSearch}\" OR publication LIKE[c] \"${fuzzyFormatedSearch}\" OR note LIKE[c] \"${fuzzyFormatedSearch}\")`,
        );
      } else if (this.searchMode === "advanced") {
        // Replace comparison operators for 'addTime'
        const compareDateMatch = formatedSearch.match(
          /(<|<=|>|>=)\s*\[\d+ DAYS\]/g,
        );
        if (compareDateMatch) {
          for (const match of compareDateMatch) {
            if (formatedSearch.includes("<")) {
              formatedSearch = formatedSearch.replaceAll(
                match,
                match.replaceAll("<", ">"),
              );
            } else if (formatedSearch.includes(">")) {
              formatedSearch = formatedSearch.replaceAll(
                match,
                match.replaceAll(">", "<"),
              );
            }
          }
        }

        // Replace Date string
        const dateRegex = /\[\d+ DAYS\]/g;
        const dateMatch = formatedSearch.match(dateRegex);
        if (dateMatch) {
          const date = new Date();
          // replace with date like: 2021-02-20@17:30:15:00
          date.setDate(date.getDate() - parseInt(dateMatch[0].slice(1, -6)));
          formatedSearch = formatedSearch.replace(
            dateRegex,
            date.toISOString().slice(0, -5).replace("T", "@"),
          );
        }
        this.filters.push(formatedSearch);
      } else if (this.searchMode === "fulltext") {
        this.filters.push(`(fulltext contains[c] \"${formatedSearch}\")`);
      }
    }

    if (this.flaged) {
      this.filters.push(`(flag == true)`);
    }
    if (this.tag) {
      this.filters.push(`(ANY tags.name == \"${this.tag}\")`);
    }
    if (this.folder) {
      this.filters.push(`(ANY folders.name == \"${this.folder}\")`);
    }
  }

  toString() {
    const filterStr = this.filters.join(" AND ");
    if (this.limit) {
      return `${filterStr} LIMIT(${this.limit})`;
    } else {
      return filterStr;
    }
  }
}
